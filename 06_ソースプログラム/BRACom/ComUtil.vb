Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者                   | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2019/05/15 | Daiko                    | 新規作成
'| 002 | 2020/11/24 | TSP                      | 追加要件No15 修正
'| 003 | 2020/11/27 | TSP                      | 要件No1対応
'| 004 | 2020/03/17 | TSP                      | 受入テスト指摘修正
'| 005 | 2020/03/17 | 日本コンピュータシステム | 要件No.1-③対応
'| 006 | 2021/12/17 | 日本コンピュータシステム | 要件No9対応
'| 007 | 2021/12/23 | 日本コンピュータシステム | 要件No.10対応
'| 008 | 2021/12/23 | 日本コンピュータシステム | 要件No.11対応
'| 009 | 2022/01/18 | 日本コンピュータシステム | 要件No.2対応
'| 010 | 2022/01/21 | 日本コンピュータシステム | 要件No.3対応
'| 011 | 2022/10/11 | Daiko                    | 要件No1 バージョン区分追加
'| 012 | 2022/11/21 | Daiko                    | 要件No1 制度受取金・積立金等項目名称設定処理を追加
'| 013 | 2022/12/15 | Daiko                    | 要件No4 バージョン区分追加
'| 014 | 2022/12/20 | Daiko                    | 要件No4 集計結果検討表（報告用）追加
'| 015 | 2023/01/11 | Daiko                    | 要件No.4 制度受取金・積立金等項目名称取得処理追加
'| 016 | 2023/01/25 | Daiko                    | 要件No.15 集計結果検討表（報告用）追加
'| 017 | 2023/03/14 | Daiko                    | 要件No.12 牛乳の生産費平均値種類が0以外の場合はNo3シートを非表示化
'| 018 | 2023/01/25 | Daiko                    | 要件No.19対応
'| 019 | 2023/04/28 | Daiko                    | 変更要件No.3対応
'| 020 | 2023/08/07 | Daiko                    | 要件No.3-② 子牛生産費の対象畜概要１の個体識別番号が空の行を削除
'| 021 | 2023/08/07 | Daiko                    | 要件No.13-② 計算ボタン削除
'| 022 | 2023/08/08 | Daiko                    | 要件No.8 "前"付の項番の存在チェックを許容
'| 023 | 2024/05/31 | Daiko                    | 要件No.1,追加要件No.3
'| 024 | 2025/09/11 | GCU                      | 要件No.12
'| 025 | 2025/08/28 | GCU                      | 要件No.2 継続区分追加
'------------------------------------------------------------------------------------------

''' <summary>
''' 共通処理
''' </summary>
''' <remarks></remarks>
Public Class ComUtil

    'REV-005 START-------------------
    ''' <summary>調査票項目マスタ呼び出し時の引数用（調査年）</summary>
    Private Const 調査年_引数用 As String = "2023"
    'REV-005 END-------------------

    ''' <summary>
    ''' DataGridViewのプロパティを設定する
    ''' </summary>
    ''' <param name="dgv"></param>
    Public Shared Sub ConfigDgv(ByVal dgv As DataGridView)
        '列ヘッダーの高さを変更できないようにする
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        '行ヘッダーの幅を変更できないようにする
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        '列の幅を変更できないようにする
        dgv.AllowUserToResizeColumns = False
        '行の高さを変更できないようにする
        dgv.AllowUserToResizeRows = False
        '行を追加できないようにする
        dgv.AllowUserToAddRows = False
        '行を削除できないようにする
        dgv.AllowUserToDeleteRows = False
        '複数セルの選択をできないようにする
        dgv.MultiSelect = False
        '行ヘッダーを含んでいる列を非表示にする
        dgv.RowHeadersVisible = False
        '列ヘッダーのテキストを中央揃えにする
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '列単位の設定を行う
        For Each dgvCol As DataGridViewColumn In dgv.Columns
            'ソートできないようにする
            dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable
            '列のサイズを変更できないようにする
            dgvCol.Resizable = DataGridViewTriState.False
            'チェックボックスの場合
            If TypeOf dgvCol Is DataGridViewCheckBoxColumn Then
                '列のセルを編集できるようにする
                dgvCol.ReadOnly = False
            Else
                '列のセルを編集できないようにする
                dgvCol.ReadOnly = True
            End If
        Next
    End Sub

    ''' <summary>
    ''' 編集可DataGridViewのプロパティを設定する
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="row"></param>
    Public Shared Sub ConfigDgvEditable(ByVal dgv As DataGridView, row As Integer)
        '列ヘッダーの高さを変更できないようにする
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        '行ヘッダーの幅を変更できないようにする
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        '列の幅を変更できないようにする
        dgv.AllowUserToResizeColumns = False
        '行の高さを変更できないようにする
        dgv.AllowUserToResizeRows = False
        '行を追加できないようにする
        dgv.AllowUserToAddRows = False
        '行を削除できないようにする
        dgv.AllowUserToDeleteRows = False
        '複数セルの選択をできないようにする
        dgv.MultiSelect = False
        '行ヘッダーを含んでいる列を非表示にする
        dgv.RowHeadersVisible = False
        '列ヘッダーのテキストを中央揃えにする
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '列単位の設定を行う
        For Each dgvCol As DataGridViewColumn In dgv.Columns
            'ソートできないようにする
            dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable
            '列のサイズを変更できないようにする
            dgvCol.Resizable = DataGridViewTriState.False
        Next

        '行の追加
        For i As Integer = 1 To row
            dgv.Rows.Add()
        Next
    End Sub


    ''' <summary>
    ''' DataGridViewチェック列全設定
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="chk"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetDataGridViewAllCheck(dgv As DataGridView, chk As Boolean)
        For i As Integer = 0 To dgv.Rows.Count - 1
            dgv(0, i).Value = chk
        Next
    End Sub

    ''' <summary>
    ''' DataGridViewチェック列全設定（活性列のみ）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="chk"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetDataGridViewAllCheckEnabledOnly(dgv As DataGridView, chk As Boolean)
        For i As Integer = 0 To dgv.Rows.Count - 1
            If dgv(0, i).ReadOnly = False Then
                dgv(0, i).Value = chk
            End If
        Next
    End Sub

    ''' <summary>
    ''' 区分リスト１取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetKubun1List(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.区分１.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.区分１.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 区分リスト２取得
    ''' </summary>
    ''' <param name="cboKubun1"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetKubun2List(cboKubun1 As String, Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.区分２.リスト.Keys
            If ComConst.区分２.リスト(Key).区分１.Contains(cboKubun1) Then
                lst.Add(New DictionaryEntry(Key, ComConst.区分２.リスト(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 調査区分リスト取得
    ''' </summary>
    ''' <param name="cboKubun1"></param>
    ''' <param name="cboKubun2"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetChosaKubunList(cboKubun1 As String, cboKubun2 As String, Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.調査区分.リスト.Keys
            If ComConst.調査区分.リスト(Key).区分１ = cboKubun1 And ComConst.調査区分.リスト(Key).区分２ = cboKubun2 Then
                lst.Add(New DictionaryEntry(Key, ComConst.調査区分.リスト(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 局リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetKyokuList(Optional none As Boolean = False) As DataTable
        Dim ret As DataTable = MasterDao.GetKyokuMaster()

        If none Then
            Dim row As DataRow = ret.NewRow
            ret.Rows.InsertAt(row, 0)
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 拠点リスト取得
    ''' </summary>
    ''' <param name="kyoku"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetKyotenList(kyoku As String, Optional jimusho As String = Nothing, Optional none As Boolean = False) As DataTable
        Dim ret As DataTable = MasterDao.GetCenterMaster(kyoku, jimusho)

        If none Then
            Dim row As DataRow = ret.NewRow
            ret.Rows.InsertAt(row, 0)
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 営農類型リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetEinouRuikeiList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.営農類型区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.営農類型区分.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 欠測値補完リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetKessokuchiHokanList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.欠測値補完.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.欠測値補完.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 貸借対照表区分リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetTaishakuTaishohyoList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.貸借対照表区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.貸借対照表区分.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 営農経営体区分リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetEinouKeieitaiList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.営農経営体区分.リスト.Keys
            If ComConst.営農経営体区分.リスト(Key).調査区分.Contains(CommonInfo.Chosakubun) Then
                lst.Add(New DictionaryEntry(Key, ComConst.営農経営体区分.リスト(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 平均種類リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetHeikinSyuruiList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.平均種類.リスト.Keys
            If ComConst.平均種類.リスト(Key).区分２.Contains(CommonInfo.Kubun2) Then
                lst.Add(New DictionaryEntry(Key, ComConst.平均種類.リスト(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 任意階層利用リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetNiniKaisouList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.任意階層利用.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.任意階層利用.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 規模階層リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetKiboKaisouList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.規模階層.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.規模階層.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 集計区分リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetSyukeiKubunList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.集計区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.集計区分.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 田畑区分リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetTahataKubunList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.田畑区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.田畑区分.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' ビール麦販売区分リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetBeerMugiKubunList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.ビール麦販売区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.ビール麦販売区分.リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' てんさい栽培区分リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetTensaiKubunList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.てんさい栽培区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.てんさい栽培区分.リスト(Key)))
        Next

        Return lst
    End Function

    'REV-016 START-------------------
    ''' <summary>
    ''' 集計１リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetSyuukei1List(cboEinouKeieitai As String, versionKbn As String, Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.集計１.リスト(versionKbn).Keys
            If ComConst.集計１.リスト(versionKbn)(Key).営農経営体区分.Equals(cboEinouKeieitai) Then
                lst.Add(New DictionaryEntry(Key, ComConst.集計１.リスト(versionKbn)(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 集計２リスト取得
    ''' </summary>
    ''' <param name="cboSyukei1"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetSyuukei2List(cboEinouKeieitai As String, cboSyukei1 As String, versionKbn As String, Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If
        For Each Key In ComConst.集計２.リスト(versionKbn).Keys
            If ComConst.集計２.リスト(versionKbn)(Key).営農経営体区分.Contains(cboEinouKeieitai) And ComConst.集計２.リスト(versionKbn)(Key).集計１.Contains(cboSyukei1) Then
                lst.Add(New DictionaryEntry(Key, ComConst.集計２.リスト(versionKbn)(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 集計３リスト取得
    ''' </summary>
    ''' <param name="cboSyukei1"></param>
    ''' <param name="cboSyukei2"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetSyuukei3List(cboEinouKeieitai As String, cboSyukei1 As String, cboSyukei2 As String, versionKbn As String, Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.集計３.リスト(versionKbn).Keys
            If ComConst.集計３.リスト(versionKbn)(Key).営農経営体区分.Contains(cboEinouKeieitai) And
                   ComConst.集計３.リスト(versionKbn)(Key).集計１.Contains(cboSyukei1) And
                   ComConst.集計３.リスト(versionKbn)(Key).集計２.Contains(cboSyukei2) Then
                lst.Add(New DictionaryEntry(Key, ComConst.集計３.リスト(versionKbn)(Key).名称))
            End If
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 集計４リスト取得
    ''' </summary>
    ''' <param name="cboSyukei1"></param>
    ''' <param name="cboSyukei2"></param>
    ''' <param name="cboSyukei3"></param>
    ''' <param name="versionKbn"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetSyuukei4List(cboEinouKeieitai As String, cboSyukei1 As String, cboSyukei2 As String, cboSyukei3 As String, versionKbn As String, Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.集計４.リスト(versionKbn).Keys
            If ComConst.集計４.リスト(versionKbn)(Key).営農経営体区分.Contains(cboEinouKeieitai) And
                   ComConst.集計４.リスト(versionKbn)(Key).集計１.Contains(cboSyukei1) And
                   ComConst.集計４.リスト(versionKbn)(Key).集計２.Contains(cboSyukei2) And
                   ComConst.集計４.リスト(versionKbn)(Key).集計３.Contains(cboSyukei3) Then
                lst.Add(New DictionaryEntry(Key, ComConst.集計４.リスト(versionKbn)(Key).名称))
            End If
        Next

        Return lst
    End Function
    'REV-016 END-------------------

    ''' <summary>
    ''' 地域区分リスト取得
    ''' </summary>
    ''' <param name="cboChiiki"></param>
    ''' <param name="none"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetChiikiKbnList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.地域区分.リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.地域区分.リスト(Key)))
        Next

        Return lst
    End Function

    ' REV_011↓
    ''' <summary>
    ''' バージョン区分体系リスト取得
    ''' </summary>
    ''' <param name="none"></param>
    ''' <returns></returns>
    Private Shared Function GetVersionKbnList(Optional none As Boolean = False) As ArrayList
        Dim lst As New ArrayList()

        If none Then
            lst.Add(New DictionaryEntry(Nothing, Nothing))
        End If

        For Each Key In ComConst.バージョン区分.体系リスト.Keys
            lst.Add(New DictionaryEntry(Key, ComConst.バージョン区分.体系リスト(Key)))
        Next

        Return lst
    End Function

    ''' <summary>
    ''' 専門調査員判定
    ''' </summary>
    ''' <param name="db"></param>
    ''' <param name="userID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsSenmonChosain(db As DBAccess, userID As String) As Boolean
        Dim ret As Boolean = False

        If DAOOther.CheckSenmonChosainExist(db, userID) Then
            ret = True
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 当該工程の接続文字列を取得する
    ''' </summary>
    ''' <param name="koutei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCurrentConnectionString(ByVal koutei As String) As String
        Dim ret As String = String.Empty
        '接続文字列取得
        Select Case koutei
            Case CommonInfo.KouteiKubun.Code.Center
                ret = My.Settings.BRASConnectionString
            Case CommonInfo.KouteiKubun.Code.Kyoku
                ret = My.Settings.BRANConnectionString
            Case CommonInfo.KouteiKubun.Code.Honsyo
                ret = My.Settings.BRAHConnectionString
        End Select

        Return ret
    End Function

    ''' <summary>
    ''' 区分１コンボボックス設定
    ''' </summary>
    ''' <param name="kubun1"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetKubun1ComboBox(kubun1 As ComboBox)
        kubun1.DisplayMember = "Value"
        kubun1.ValueMember = "Key"
        kubun1.DataSource = ComUtil.GetKubun1List(True)
    End Sub

    ''' <summary>
    ''' 区分２コンボボックス設定
    ''' </summary>
    ''' <param name="kubun1"></param>
    ''' <param name="kubun2"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetKubun2ComboBox(kubun1 As ComboBox, kubun2 As ComboBox)
        kubun2.DisplayMember = "Value"
        kubun2.ValueMember = "Key"

        If kubun1.SelectedValue Is Nothing Then
            kubun2.DataSource = Nothing
            kubun2.Items.Clear()
            kubun2.Items.Add("")
        Else
            kubun2.DataSource = ComUtil.GetKubun2List(kubun1.SelectedValue.ToString, True)
        End If
    End Sub

    ''' <summary>
    ''' 調査区分コンボボックス設定
    ''' </summary>
    ''' <param name="kubun1"></param>
    ''' <param name="kubun2"></param>
    ''' <param name="chosakubun"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetChosakubunComboBox(kubun1 As ComboBox, kubun2 As ComboBox, chosakubun As ComboBox)
        chosakubun.DisplayMember = "Value"
        chosakubun.ValueMember = "Key"

        If kubun2.SelectedValue Is Nothing Then
            chosakubun.DataSource = Nothing
            chosakubun.Items.Clear()
            chosakubun.Items.Add("")
        Else
            chosakubun.DataSource = ComUtil.GetChosaKubunList(kubun1.SelectedValue.ToString, kubun2.SelectedValue.ToString, True)
        End If
    End Sub

    ''' <summary>
    ''' 局コンボボックス設定
    ''' </summary>
    ''' <param name="kyoku"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetKyokuComboBox(kyoku As ComboBox)
        kyoku.DisplayMember = "局名"
        kyoku.ValueMember = "局コード"
        kyoku.DataSource = ComUtil.GetKyokuList(True)

        If Not CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
            kyoku.SelectedValue = CommonInfo.Kyoku
            kyoku.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' 拠点コンボボックス設定
    ''' </summary>
    ''' <param name="kyoku"></param>
    ''' <param name="kyoten"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetKyotenComboBox(kyoku As ComboBox, kyoten As ComboBox)
        kyoten.DisplayMember = "センター名"
        kyoten.ValueMember = "センター番号"

        If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Center Then
            kyoten.DataSource = ComUtil.GetKyotenList(kyoku.SelectedValue.ToString, CommonInfo.Jimusyo, True)
            kyoten.SelectedValue = CommonInfo.Center
            kyoten.Enabled = False
        Else
            kyoten.DataSource = ComUtil.GetKyotenList(kyoku.SelectedValue.ToString, , True)
        End If
    End Sub

    ''' <summary>
    ''' 営農類型コンボボックス設定
    ''' </summary>
    ''' <param name="lbl"></param>
    ''' <param name="cbo"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetEinouRuikeiComboBox(lbl As Label, cbo As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人) Then

            lbl.Visible = False
            cbo.Visible = False
        Else
            cbo.ValueMember = "Key"
            cbo.DisplayMember = "Value"
            cbo.DataSource = ComUtil.GetEinouRuikeiList(True)
        End If
    End Sub

    ''' <summary>
    ''' 欠測値補完コンボボックス設定
    ''' </summary>
    ''' <param name="lbl"></param>
    ''' <param name="cbo"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetKessokuchiHokanComboBox(lbl As Label, cbo As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
            And CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo) Then

            lbl.Visible = False
            cbo.Visible = False
        Else
            cbo.ValueMember = "Key"
            cbo.DisplayMember = "Value"
            cbo.DataSource = ComUtil.GetKessokuchiHokanList(False)
        End If
    End Sub

    ''' <summary>
    ''' 貸借対照表区分コンボボックス設定
    ''' </summary>
    ''' <param name="lbl"></param>
    ''' <param name="cbo"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetTaishakuTaishohyoComboBox(lbl As Label, cbo As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
            And CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo) Then

            lbl.Visible = False
            cbo.Visible = False
        Else
            cbo.ValueMember = "Key"
            cbo.DisplayMember = "Value"
            cbo.DataSource = ComUtil.GetTaishakuTaishohyoList(True)
        End If
    End Sub

    ''' <summary>
    ''' 営農経営体区分コンボボックス設定
    ''' </summary>
    ''' <param name="lbl"></param>
    ''' <param name="cbo"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetEinouKeieitaiComboBox(lbl As Label, cbo As ComboBox)
        If Not (CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計) Then

            lbl.Visible = False
            cbo.Visible = False
        Else
            cbo.ValueMember = "Key"
            cbo.DisplayMember = "Value"
            cbo.DataSource = ComUtil.GetEinouKeieitaiList(False)
        End If
    End Sub

    ''' <summary>
    ''' 平均種類コンボボックス設定
    ''' </summary>
    ''' <param name="heikinSyurui"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetHeikinSyuruiComboBox(heikinSyurui As ComboBox)
        heikinSyurui.DisplayMember = "Value"
        heikinSyurui.ValueMember = "Key"
        heikinSyurui.DataSource = ComUtil.GetHeikinSyuruiList(False)
    End Sub

    ''' <summary>
    ''' 任意階層利用コンボボックス設定
    ''' </summary>
    ''' <param name="kiboKaisou"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetNiniKaisouComboBox(niniKaisou As ComboBox)
        niniKaisou.DisplayMember = "Value"
        niniKaisou.ValueMember = "Key"
        niniKaisou.DataSource = ComUtil.GetNiniKaisouList(False)
    End Sub

    ''' <summary>
    ''' 規模階層コンボボックス設定
    ''' </summary>
    ''' <param name="kiboKaisou"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetKiboKaisouComboBox(kiboKaisou As ComboBox)
        kiboKaisou.DisplayMember = "Value"
        kiboKaisou.ValueMember = "Key"
        kiboKaisou.DataSource = ComUtil.GetKiboKaisouList(False)
    End Sub

    'REV 025↓
    ''' <summary>
    ''' 継続区分コンボボックス設定
    ''' </summary>
    ''' <param name="lblKeizokuKubun">継続区分ラベル</param>
    ''' <param name="cboKeizokuKubun">継続区分コンボボックス</param>
    ''' <param name="chkZenkaiCensus">前回センサス番号を使用チェックボックス</param>
    ''' <param name="cboEinouKeieitai">営農経営体区分コンボボックス</param>
    ''' <remarks></remarks>
    Public Shared Sub SetKeizokuKubunComboBox(lblKeizokuKubun As Label, cboKeizokuKubun As ComboBox, chkZenkaiCensus As CheckBox, cboEinouKeieitai As ComboBox)
        Dim lst As New ArrayList()
        lst.Add(New DictionaryEntry("1", "全て"))
        lst.Add(New DictionaryEntry("2", "継続のみ"))

        cboKeizokuKubun.DisplayMember = "Value"
        cboKeizokuKubun.ValueMember = "Key"
        cboKeizokuKubun.DataSource = lst

        ' 初期表示条件設定
        SetKeizokuKubunVisibility(lblKeizokuKubun, cboKeizokuKubun, chkZenkaiCensus, cboEinouKeieitai)

        ' 初期活性状態設定
        SetZenkaiCensusEnabled(chkZenkaiCensus, cboKeizokuKubun)
    End Sub

    ''' <summary>
    ''' 継続区分の表示制御
    ''' </summary>
    ''' <param name="lblKeizokuKubun">継続区分ラベル</param>
    ''' <param name="cboKeizokuKubun">継続区分コンボボックス</param>
    ''' <param name="chkZenkaiCensus">前回センサス番号を使用チェックボックス</param>
    ''' <param name="cboEinouKeieitai">営農経営体区分コンボボックス</param>
    ''' <remarks></remarks>
    Public Shared Sub SetKeizokuKubunVisibility(lblKeizokuKubun As Label, cboKeizokuKubun As ComboBox, chkZenkaiCensus As CheckBox, cboEinouKeieitai As ComboBox)
        Dim chosakubun As String = ComUtil.GetChosakubun(cboEinouKeieitai)

        ' 営農類型別経営統計（個人・法人）の場合は非表示、それ以外は表示
        If chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Or chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            lblKeizokuKubun.Visible = False
            cboKeizokuKubun.Visible = False
        Else
            lblKeizokuKubun.Visible = True
            cboKeizokuKubun.Visible = True
        End If

        ' チェックボックスは継続区分と同じ表示制御
        chkZenkaiCensus.Visible = lblKeizokuKubun.Visible
    End Sub

    ''' <summary>
    ''' 前回センサス番号使用チェックボックスの活性制御
    ''' </summary>
    ''' <param name="chkZenkaiCensus">前回センサス番号を使用チェックボックス</param>
    ''' <param name="cboKeizokuKubun">継続区分コンボボックス</param>
    ''' <remarks></remarks>
    Public Shared Sub SetZenkaiCensusEnabled(chkZenkaiCensus As CheckBox, cboKeizokuKubun As ComboBox)
        If cboKeizokuKubun.SelectedValue IsNot Nothing Then
            If cboKeizokuKubun.SelectedValue.ToString() = ComConst.継続区分.継続のみ Then
                ' 継続のみの場合は活性
                chkZenkaiCensus.Enabled = True
            Else
                ' 全ての場合は非活性
                chkZenkaiCensus.Enabled = False
                chkZenkaiCensus.Checked = False
            End If
        End If
    End Sub
    'REV 025↑

    ''' <summary>
    ''' 集計区分コンボボックス設定
    ''' </summary>
    ''' <param name="syukeiKubun"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSyukeiKubunComboBox(syukeiKubun As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.六条大麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.はだか麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.そば生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.大豆生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.原料用かんしょ生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.原料用ばれいしょ生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.なたね生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.てんさい生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.さとうきび生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_組織法人 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.大豆生産費統計_組織法人 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_六条大麦生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_はだか麦生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_そば生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_なたね生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_てんさい生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_さとうきび生産費) Then

            syukeiKubun.Enabled = False
        Else
            syukeiKubun.DisplayMember = "Value"
            syukeiKubun.ValueMember = "Key"
            syukeiKubun.DataSource = ComUtil.GetSyukeiKubunList(False)
        End If
    End Sub

    ''' <summary>
    ''' 田畑区分コンボボックス設定
    ''' </summary>
    ''' <param name="tahataKubun"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetTahataKubunComboBox(tahataKubun As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.小麦生産費統計_組織法人 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.六条大麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_六条大麦生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.はだか麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_はだか麦生産費 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.大豆生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.大豆生産費統計_組織法人) Then

            tahataKubun.Enabled = False
        Else
            tahataKubun.DisplayMember = "Value"
            tahataKubun.ValueMember = "Key"
            tahataKubun.DataSource = ComUtil.GetTahataKubunList(False)
        End If
    End Sub

    ''' <summary>
    ''' ビール麦販売区分コンボボックス設定
    ''' </summary>
    ''' <param name="beerMugiKubun"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetBeerMugiKubunComboBox(beerMugiKubun As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費) Then

            beerMugiKubun.Enabled = False
        Else
            beerMugiKubun.DisplayMember = "Value"
            beerMugiKubun.ValueMember = "Key"
            beerMugiKubun.DataSource = ComUtil.GetBeerMugiKubunList(False)
        End If
    End Sub

    ''' <summary>
    ''' てんさい栽培区分コンボボックス設定
    ''' </summary>
    ''' <param name="tensaiKubun"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetTensaiKubunComboBox(tensaiKubun As ComboBox)
        If Not (CommonInfo.Chosakubun = ComConst.調査区分.てんさい生産費統計_個別 _
            Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_てんさい生産費) Then

            tensaiKubun.Enabled = False
        Else
            tensaiKubun.DisplayMember = "Value"
            tensaiKubun.ValueMember = "Key"
            tensaiKubun.DataSource = ComUtil.GetTensaiKubunList(False)
        End If
    End Sub

    'REV-016 START-------------------
    ''' <summary>
    ''' 集計１コンボボックス設定
    ''' </summary>
    ''' <param name="syukei1"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSyukei1ComboBox(einouKeieitai As ComboBox, syukei1 As ComboBox, versionKbn As String)
        syukei1.DisplayMember = "Value"
        syukei1.ValueMember = "Key"
        syukei1.DataSource = ComUtil.GetSyuukei1List(einouKeieitai.SelectedValue.ToString, versionKbn, True)
        changeDropdownWidth(syukei1)
    End Sub

    ''' <summary>
    ''' 集計２コンボボックス設定
    ''' </summary>
    ''' <param name="syukei1"></param>
    ''' <param name="syukei2"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSyukei2ComboBox(einouKeieitai As ComboBox, syukei1 As ComboBox, syukei2 As ComboBox, versionKbn As String)

        If syukei1.SelectedValue Is Nothing Then
            syukei2.DataSource = Nothing
            syukei2.Items.Clear()
            syukei2.Items.Add("")
        Else
            syukei2.DisplayMember = "Value"
            syukei2.ValueMember = "Key"
            syukei2.DataSource = ComUtil.GetSyuukei2List(einouKeieitai.SelectedValue.ToString, syukei1.SelectedValue.ToString, versionKbn, True)
            changeDropdownWidth(syukei2)
        End If
    End Sub

    ''' <summary>
    ''' 集計３コンボボックス設定
    ''' </summary>
    ''' <param name="syukei1"></param>
    ''' <param name="syukei2"></param>
    ''' <param name="syukei3"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSyukei3ComboBox(einouKeieitai As ComboBox, syukei1 As ComboBox, syukei2 As ComboBox, syukei3 As ComboBox, versionKbn As String)
        syukei3.DisplayMember = "Value"
        syukei3.ValueMember = "Key"
        If syukei2.SelectedValue Is Nothing Then
            syukei3.DataSource = Nothing
            syukei3.Items.Clear()
            syukei3.Items.Add("")
        Else
            syukei3.DataSource = ComUtil.GetSyuukei3List(einouKeieitai.SelectedValue.ToString, syukei1.SelectedValue.ToString, syukei2.SelectedValue.ToString, versionKbn, True)
            changeDropdownWidth(syukei3)
        End If
    End Sub

    ''' <summary>
    ''' 集計４コンボボックス設定
    ''' </summary>
    ''' <param name="syukei1"></param>
    ''' <param name="syukei2"></param>
    ''' <param name="syukei3"></param>
    ''' <param name="syukei4"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSyukei4ComboBox(einouKeieitai As ComboBox, syukei1 As ComboBox, syukei2 As ComboBox, syukei3 As ComboBox, syukei4 As ComboBox, versionKbn As String)
        syukei4.DisplayMember = "Value"
        syukei4.ValueMember = "Key"
        If syukei3.SelectedValue Is Nothing Then
            syukei4.DataSource = Nothing
            syukei4.Items.Clear()
            syukei4.Items.Add("")
        Else
            syukei4.DataSource = ComUtil.GetSyuukei4List(einouKeieitai.SelectedValue.ToString, syukei1.SelectedValue.ToString, syukei2.SelectedValue.ToString, syukei3.SelectedValue.ToString, versionKbn, True)
            changeDropdownWidth(syukei4)
        End If
    End Sub
    'REV-016 END-------------------

    ''' <summary>
    ''' 地域区分コンボボックス設定
    ''' </summary>
    ''' <param name="chiikiKbn"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetChiikiKbnComboBox(chiikiKbn As ComboBox)
        chiikiKbn.DisplayMember = "Value"
        chiikiKbn.ValueMember = "Key"
        chiikiKbn.DataSource = ComUtil.GetChiikiKbnList(False)
    End Sub

    ' REV_011↓
    ''' <summary>
    ''' バージョン区分コンボボックス設定
    ''' </summary>
    ''' <param name="versionKbn"></param>
    Public Shared Sub SetVersionKbnComboBox(versionKbn As ComboBox)
        versionKbn.DisplayMember = "Value"
        versionKbn.ValueMember = "Key"
        versionKbn.DataSource = ComUtil.GetVersionKbnList(False)
    End Sub
    ' REV_011↑

    ''' <summary>
    ''' ドロップダウンリストの幅を一番長い文字列に合わせる
    ''' </summary>
    ''' <param name="cbo"></param>
    ''' <remarks></remarks>
    Private Shared Sub changeDropdownWidth(cbo As ComboBox)
        Dim maxSize As Integer = 0
        For Each item As DictionaryEntry In cbo.Items
            maxSize = Math.Max(maxSize, TextRenderer.MeasureText(CStr(item.Value), cbo.Font).Width)
        Next

        If Not maxSize = 0 Then
            cbo.DropDownWidth = maxSize
        Else
            cbo.DropDownWidth = cbo.Width
        End If
    End Sub

    ''' <summary>
    ''' 事務所番号変換（北海道対応）
    ''' </summary>
    ''' <param name="todofuken"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvJimusyoNo(todofuken As String) As String
        Dim ret As String = String.Empty
        Dim val As Integer

        If Integer.TryParse(todofuken, val) Then
            If val = 1 Then
                val = 51
                ret = val.ToString
            Else
                ret = todofuken
            End If
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 都道府県番号変換（北海道対応）
    ''' </summary>
    ''' <param name="Jimusyo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvTodofukenNo(Jimusyo As String) As String
        Dim ret As String = String.Empty
        Dim val As Integer

        If Integer.TryParse(Jimusyo, val) Then
            If val = 51 Then
                val = 1
                ret = val.ToString
            Else
                ret = Jimusyo
            End If
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 都道府県取得
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTodofuken(censusNo As String) As String
        Return censusNo.Substring(0, 2)
    End Function

    ''' <summary>
    ''' 市区町村取得
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetShikuchoson(censusNo As String) As String
        Return censusNo.Substring(2, 3)
    End Function

    ''' <summary>
    ''' 旧市区町村取得
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKyuShikuchoson(censusNo As String) As String
        Return censusNo.Substring(5, 2)
    End Function

    ''' <summary>
    ''' 農業集落取得
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNogyoShuraku(censusNo As String) As String
        Return censusNo.Substring(7, 3)
    End Function

    ''' <summary>
    ''' 調査区取得
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosaku(censusNo As String) As String
        Return censusNo.Substring(10, 3)
    End Function

    ''' <summary>
    ''' 客体番号取得
    ''' </summary>
    ''' <param name="censusNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKyakutaiNo(censusNo As String) As String
        Return censusNo.Substring(13, 3)
    End Function

    ''' <summary>
    ''' センサス番号取得
    ''' </summary>
    ''' <param name="todofuken"></param>
    ''' <param name="shikuchoson"></param>
    ''' <param name="kyuShikuchoson"></param>
    ''' <param name="nogyoShuraku"></param>
    ''' <param name="chosaku"></param>
    ''' <param name="kyakutaiNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCensusNo(todofuken As String, shikuchoson As String, kyuShikuchoson As String, nogyoShuraku As String, chosaku As String, kyakutaiNo As String) As String
        Return todofuken.PadLeft(2, "0"c) _
             & shikuchoson.PadLeft(3, "0"c) _
             & kyuShikuchoson.PadLeft(2, "0"c) _
             & nogyoShuraku.PadLeft(3, "0"c) _
             & chosaku.PadLeft(3, "0"c) _
             & kyakutaiNo.PadLeft(3, "0"c)
    End Function

    ''' <summary>
    ''' フォルダパス取得
    ''' </summary>
    ''' <param name="selectedPath"></param>
    ''' <param name="frm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFolderPath(frm As Form, selectedPath As String) As String
        Dim ret As String = String.Empty
        Dim fbd As New FolderBrowserDialog

        fbd.Description = "フォルダを指定してください。"
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        fbd.SelectedPath = selectedPath
        fbd.ShowNewFolderButton = False

        If Not System.IO.Directory.Exists(selectedPath) Then
            System.IO.Directory.CreateDirectory(selectedPath)
        End If

        If fbd.ShowDialog(frm) = DialogResult.OK Then
            ret = fbd.SelectedPath
        End If

        Return ret
    End Function

    ''' <summary>
    ''' ファイルパス取得
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="frm"></param>
    ''' <param name="selectedPath"></param>
    ''' <param name="fileName"></param>
    ''' <param name="filter"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFilePath(Of T As {FileDialog, New})(frm As Form, selectedPath As String, Optional fileName As String = Nothing, Optional filter As String = "Excelファイル (*.xlsx)|*.xlsx") As String
        Dim ret As String = String.Empty
        Dim fdg As New T

        fdg.InitialDirectory = selectedPath
        fdg.Filter = filter
        fdg.FilterIndex = 1
        fdg.RestoreDirectory = True
        If Not fileName Is Nothing Then
            fdg.FileName = fileName
        End If
        If TypeOf fdg Is SaveFileDialog Then
            Dim tp As Type = fdg.GetType()
            Dim pInf As Reflection.PropertyInfo = tp.GetProperty("OverwritePrompt")
            pInf.SetValue(fdg, False)
        End If

        If Not System.IO.Directory.Exists(selectedPath) Then
            System.IO.Directory.CreateDirectory(selectedPath)
        End If

        If fdg.ShowDialog(frm) = DialogResult.OK Then
            ret = fdg.FileName
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 調査区分取得
    ''' </summary>
    ''' <param name="einouKeieitai"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosakubun(einouKeieitai As ComboBox) As String
        Dim ret As String

        '営農経営体区分判定取得
        ret = If(einouKeieitai.SelectedValue Is Nothing, CommonInfo.Chosakubun, einouKeieitai.SelectedValue.ToString)

        Return ret
    End Function

    ''' <summary>
    ''' 調査区分名取得
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosakubunName(chosakubun As String) As String
        Dim ret As String

        ret = If(chosakubun.Equals(ComConst.営農経営体区分.農業経営体), ComConst.営農経営体区分.リスト(ComConst.営農経営体区分.農業経営体).名称２, CommonInfo.ChosakubunName)

        Return ret
    End Function

    ''' <summary>
    ''' 調査区分取得
    ''' </summary>
    ''' <param name="einouKeieitai"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChosakubun(einouKeieitai As String) As String
        Dim ret As String

        '営農経営体区分判定取得
        ret = If(einouKeieitai Is Nothing, CommonInfo.Chosakubun, einouKeieitai)

        Return ret
    End Function

    ''' <summary>
    ''' 欠測値補完取得
    ''' </summary>
    ''' <param name="kessokuchiHokana"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetkessokuchiHokana(kessokuchiHokana As ComboBox) As String
        Dim ret As String

        '欠測値補完判定取得
        ret = If(kessokuchiHokana.SelectedValue Is Nothing, ComConst.欠測値補完.無, kessokuchiHokana.SelectedValue.ToString)

        Return ret
    End Function

    ''' <summary>
    ''' Excelの列番号を列英字に変換する
    ''' </summary>
    ''' <param name="col"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ConvertColNoToLetter(ByVal col As Integer) As String
        Dim ret As String = Nothing
        Dim alpha As Integer
        Dim remainder As Integer
        alpha = CInt(Math.Truncate((col - 1) / 26))
        remainder = col - (alpha * 26)
        If alpha > 0 Then
            ret = Chr(alpha + 64)
        End If
        If remainder > 0 Then
            ret = ret & Chr(remainder + 64)
        End If
        Return ret
    End Function

    ''' <summary>COMオブジェクト解放デリゲート</summary>
    Private Delegate Sub ReleaseComObjectDelegate(Of T As Class)(ByRef pObjCom As T)

    ''' <summary>
    ''' COMオブジェクト解放
    ''' </summary>
    ''' <param name="ComObjProcess"></param>
    ''' <param name="pObjCom"></param>
    ''' <remarks></remarks>
    Private Shared Sub ReleaseComObject(ComObjProcess As ComObjectProcess, pObjCom As Object)
        'デリゲートの実行
        Dim dlg As ReleaseComObjectDelegate(Of Object) = AddressOf ComObjProcess.ReleaseComObject
        dlg.Invoke(pObjCom)
    End Sub

    ''' <summary>行の高さを自動調整デリゲート</summary>
    Private Delegate Sub SetAutoFitDelegate(rng As Excel.Range)

    ''' <summary>
    ''' 行の高さを自動調整
    ''' </summary>
    ''' <param name="xlsProcess"></param>
    ''' <param name="rng"></param>
    ''' <remarks></remarks>
    Private Shared Sub SetAutoFit(xlsProcess As ExcelProcess, rng As Excel.Range)
        'デリゲートの実行
        Dim dlg As SetAutoFitDelegate = AddressOf xlsProcess.SetAutoFit
        dlg.Invoke(rng)
    End Sub

    ''' <summary>
    ''' 調査票クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Chosahyo

        ''' <summary>
        ''' 調査年コンボボックス設定
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="db"></param>
        ''' <param name="koutei"></param>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <param name="none"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetChosaNenComboBox(chosaNen As ComboBox, db As DBAccess, koutei As String, kyoku As String, jimusho As String, kyoten As String, Optional none As Boolean = False)
            Dim dt As DataTable = Nothing
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOChosahyo.GetChosaNen(db, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOChosahyo.GetChosaNen(db, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOChosahyo.GetChosaNen(db, Nothing, Nothing, Nothing)
            End Select

            If none Then
                Dim row As DataRow = dt.NewRow
                row("調査年") = DBNull.Value
                dt.Rows.InsertAt(row, 0)
            End If

            chosaNen.ValueMember = "調査年"
            chosaNen.DisplayMember = "調査年"
            chosaNen.DataSource = dt
        End Sub

        ''' <summary>
        ''' 調査票項目取得
        ''' </summary>
        ''' <param name="dtItem"></param>
        ''' <param name="dc"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetItem(dtItem As DataTable, dc As Dictionary(Of String, DataTable)) As Dictionary(Of String, DAOChosahyo.調査票項目)
            Dim ret As New Dictionary(Of String, DAOChosahyo.調査票項目)

            Dim lstItem As List(Of DataRow) = dtItem.AsEnumerable().ToList()

            For Each kv As KeyValuePair(Of String, DataTable) In dc
                Dim tableName As String = kv.Key
                Dim dt As DataTable = kv.Value
                Dim lstDr As List(Of DataRow) = dt.AsEnumerable().ToList()
                Dim colArr(dt.Columns.Count - 1) As DataColumn
                dt.Columns.CopyTo(colArr, 0)

                Dim seidouketori_count As Integer = 9010001 'REV-010 ADD

                If Not tableName.Contains("＿可変") Then
                    For Each row As DataRow In lstDr
                        For Each col As DataColumn In colArr
                            Dim query = (From dr In lstItem Where dr("項目番号").ToString = col.ColumnName Select dr).Take(1).ToArray

                            'REV-010 ADD START---------------------
                            Dim query2 = (From dr In lstItem Where dr("項目番号").ToString = "Q" + seidouketori_count.ToString("00000000") Select dr).Take(1).ToArray
                            'REV-010 ADD END-----------------------

                            If query.Any() Then
                                Dim item As New DAOChosahyo.調査票項目
                                With item
                                    .シート名 = query(0)("シート名").ToString
                                    .行位置 = Integer.Parse(query(0)("行位置").ToString)
                                    .列位置 = Integer.Parse(query(0)("列位置").ToString)
                                    .値 = row(col.ColumnName).ToString
                                End With
                                '2022/1/27 ADD START 同じ項目番号のレコードを追加しないよう条件追加
                                If Not ret.ContainsKey(query(0)("項目番号").ToString) Then
                                    ret.Add(col.ColumnName, item)
                                End If
                                '2022/1/27 ADD END
                            End If

                            'REV-010 ADD START---------------------
                            If query2.Any() Then
                                If query2(0)("出力項目名").ToString <> "" Then
                                    Dim item2 As New DAOChosahyo.調査票項目
                                    With item2

                                        .シート名 = query2(0)("シート名").ToString
                                        .行位置 = Integer.Parse(query2(0)("行位置").ToString)
                                        .列位置 = Integer.Parse(query2(0)("列位置").ToString)
                                        .値 = query2(0)("出力項目名").ToString

                                    End With
                                    If Not ret.ContainsKey(query2(0)("項目番号").ToString) Then
                                        ret.Add(query2(0)("項目番号").ToString, item2)
                                    End If
                                End If

                            End If

                            If seidouketori_count = 9010076 Then
                                seidouketori_count = 0
                            End If

                            seidouketori_count = seidouketori_count + 1
                            'REV-010 ADD END-----------------------
                        Next
                    Next
                Else
                    Dim itemNoArr = (From dr In lstDr Group dr By dr!項目番号 Into Group Select 項目番号).ToArray
                    For Each itemNo As String In itemNoArr
                        Dim query = (From dr In lstItem Where dr("項目番号").ToString = itemNo Select dr).Take(1).ToArray
                        If query.Any() Then
                            Dim rowArr = (From dr In lstDr Where dr("項目番号").ToString = itemNo Select dr).ToArray
                            For Each row As DataRow In rowArr
                                Dim i As Integer = Integer.Parse(row("明細番号").ToString)
                                Dim increment As Integer = Integer.Parse(query(0)("可変増量").ToString)
                                Dim item As New DAOChosahyo.調査票項目
                                With item
                                    .シート名 = query(0)("シート名").ToString
                                    .行位置 = 1 + If(query(0)("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                    .列位置 = 1 + If(query(0)("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                    .値 = row("値").ToString

                                    Dim rw As Integer = Integer.Parse(query(0)("行位置").ToString)
                                    Dim cl As Integer = Integer.Parse(query(0)("列位置").ToString)
                                    Dim max As Integer = Integer.Parse(query(0)("可変最大数").ToString)
                                    Dim colLetter As String = ComUtil.ConvertColNoToLetter(cl)
                                    If query(0)("可変方向").ToString = ComConst.可変方向.縦 Then
                                        .可変範囲 = colLetter & rw & ":" & colLetter & (rw + max * increment - 1)
                                    Else
                                        .可変範囲 = colLetter & rw & ":" & ComUtil.ConvertColNoToLetter(cl + max * increment - 1) & rw
                                    End If
                                End With
                                ret.Add(row("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & row("明細番号").ToString, item)
                            Next
                        End If
                    Next
                End If
            Next

            Return ret
        End Function

        ''' <summary>
        ''' 調査票シートデータ取得
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess, _chosaNen As String) As Dictionary(Of String, DAOChosahyo.調査票項目)
            Dim ret As New Dictionary(Of String, DAOChosahyo.調査票項目)


            Dim sheets = From dr In dt Group dr By dr!シート名 Into Group Select シート名
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                    'シート保護確認
                    If xlSheet.ProtectContents Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object
                        'REV-005 START-----------------------
                        rng = xlSheet.Range(ComConst.調査票.シートデータ範囲(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun)))(sheet.ToString))
                        'REV-005 END-----------------------

                        arrData = DirectCast(rng.Value, Object(,))

                        Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr

                        '2022/1/27 ADD START 営農法人の"11_02_制度受取金・積立金"シートのみ項番をずらすよう修正
                        Dim seidouketori_count As Integer = 9010001

                        If sheet.Equals("11_02_制度受取金・積立金") Then

                            seidouketori_count = 9010031

                        End If
                        'REV-010 END

                        For Each dr As DataRow In query
                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                Dim item As New DAOChosahyo.調査票項目
                                With item
                                    'REV-010 START 出力項目名のみ保存対象から外す-----------------------
                                    If dr("項目番号").ToString = "Q" + seidouketori_count.ToString("00000000") Then

                                        seidouketori_count = seidouketori_count + 1

                                        Continue For
                                    End If
                                    'REV-010 END  -----------------------

                                    .シート名 = dr("シート名").ToString
                                    .行位置 = Integer.Parse(dr("行位置").ToString)

                                    .列位置 = Integer.Parse(dr("列位置").ToString)


                                    'REV-005 START-----------------------
                                    If .行位置 = 0 Or .列位置 = 0 Then
                                        .値 = Nothing
                                    Else
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        If .値 IsNot Nothing Then
                                            EscapeToiawasesaki(dr, .値)
                                        End If
                                    End If
                                    'REV-005 END-----------------------

                                    .型区分 = dr("型区分").ToString
                                    .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                    .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                                End With
                                ret.Add(dr("項目番号").ToString, item)
                            Else
                                Dim increment As Integer = Integer.Parse(dr("可変増量").ToString)
                                For i As Integer = 1 To Integer.Parse(dr("可変最大数").ToString)
                                    Dim item As New DAOChosahyo.調査票項目
                                    With item
                                        .シート名 = dr("シート名").ToString
                                        .行位置 = Integer.Parse(dr("行位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                        .列位置 = Integer.Parse(dr("列位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        .型区分 = dr("型区分").ToString
                                    End With
                                    If Not String.IsNullOrEmpty(item.値) Then
                                        ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i.ToString, item)
                                    End If
                                Next
                            End If
                        Next
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try

                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next

            Return ret
        End Function

        'REV-006 START-----------------------
        ''' <summary>
        ''' 【問合せ先】内容の改行をエスケープ文字に変換
        ''' </summary>
        ''' <param name="dr"></param>
        ''' <param name="値"></param>
        ''' <remarks></remarks>
        Public Shared Sub EscapeToiawasesaki(ByVal dr As DataRow, ByRef 値 As String)

            If dr("項目番号").ToString = "Q00000201" Or
               dr("項目番号").ToString = "Q00000301" Or
               dr("項目番号").ToString = "Q00000401" Then

                値 = 値.Replace(Chr(10), "\r\n")
            End If
        End Sub

        ''' <summary>
        ''' 【問合せ先】内容のエスケープ文字を改行に変換
        ''' </summary>
        ''' <param name="kv"></param>
        ''' <remarks></remarks>
        Public Shared Sub RestoreToiawasesaki(ByRef kv As KeyValuePair(Of String, DAOChosahyo.調査票項目))

            If kv.Key = "Q00000201" Or
               kv.Key = "Q00000301" Or
               kv.Key = "Q00000401" Then

                kv.Value.値 = kv.Value.値.Replace("\r\n", Chr(10))
            End If
        End Sub
        'REV-006 END-----------------------

        ''' <summary>
        ''' 調査票シートデータ設定
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dc As Dictionary(Of String, DAOChosahyo.調査票項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess)

            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    '非可変項目
                    Dim cel As Excel.Range = Nothing
                    Try
                        cel = xlSheet.Cells

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 Is Nothing Select dr
                        For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query

                            RestoreToiawasesaki(kv)
                            Dim rng As Excel.Range = Nothing
                            Try

                                'REV-005 START-----------------------
                                'ウオッチ用-----------------------
                                If kv.Value.行位置 = 0 Or kv.Value.列位置 = 0 Then
                                    kv.Value.行位置 = 0

                                Else

                                    rng = DirectCast(cel.Item(kv.Value.行位置, kv.Value.列位置), Excel.Range)
                                    rng.Value = kv.Value.値

                                End If

                            Finally
                                ReleaseComObject(comObject, rng)
                            End Try
                        Next

                    Finally
                        ReleaseComObject(comObject, cel)
                    End Try

                    '可変項目
                    Dim range As IEnumerable(Of String) = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 IsNot Nothing Select dr.Value.可変範囲 Distinct
                    For Each ar As String In range
                        Dim rng As Excel.Range = Nothing
                        Try
                            Dim arrData(,) As Object

                            rng = xlSheet.Range(ar)

                            arrData = DirectCast(rng.Formula, Object(,))

                            Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 = ar Select dr

                            For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query
                                arrData(kv.Value.行位置, kv.Value.列位置) = kv.Value.値
                            Next

                            rng.Value = arrData
                            rng.Value = rng.Formula
                        Finally
                            ReleaseComObject(comObject, rng)
                        End Try
                    Next

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 事務所番号変換（北海道対応）
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvJimusyoNo(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Dim ret As String = String.Empty
            Dim ken As String = GetTodofuken(dcChosahyo)
            Dim val As Integer

            If Integer.TryParse(ken, val) Then
                If val = 1 Then
                    val = 51
                    ret = val.ToString
                Else
                    ret = ken
                End If
            End If

            Return ret
        End Function

        ''' <summary>
        ''' 調査年取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetChosaNen(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.調査年).値
        End Function

        ''' <summary>
        ''' 都道府県取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetTodofuken(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.都道府県).値
        End Function

        ''' <summary>
        ''' 市区町村取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetShikuchoson(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.市区町村).値
        End Function

        ''' <summary>
        ''' 旧市区町村取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetKyuShikuchoson(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.旧市区町村).値
        End Function

        ''' <summary>
        ''' 農業集落取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetNogyoShuraku(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.農業集落).値
        End Function

        ''' <summary>
        ''' 調査区取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetChosaku(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.調査区).値
        End Function

        ''' <summary>
        ''' 客体番号取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetKyakutaiNo(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.客体番号).値
        End Function

        ''' <summary>
        ''' 営農類型取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetEinouRuike(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.営農類型別経営統計.営農類型).値
        End Function

        ''' <summary>
        ''' 指定品目名取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetShiteiHinmokumei(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.営農類型別経営統計.指定品目名).値
        End Function

        ''' <summary>
        ''' 対象品目取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetTaishoHinmoku(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.農産物生産費.対象品目).値
        End Function

        ''' <summary>
        ''' 経営種類取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetKeieiShurui(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.農産物生産費.経営種類).値
        End Function

        ''' <summary>
        ''' 生産費区分取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSeisanhiKubun(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Return dcChosahyo(ComConst.調査票.項目番号.畜産物生産費.生産費区分).値
        End Function

        ''' <summary>
        ''' センサス番号取得
        ''' </summary>
        ''' <param name="dcChosahyo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetCensusNo(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As String
            Dim ret As String = Nothing

            If Not (String.IsNullOrEmpty(ComUtil.Chosahyo.GetTodofuken(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetShikuchoson(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetKyuShikuchoson(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetNogyoShuraku(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetChosaku(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetKyakutaiNo(dcChosahyo))) Then

                ret = GetTodofuken(dcChosahyo).PadLeft(2, "0"c) _
                 & GetShikuchoson(dcChosahyo).PadLeft(3, "0"c) _
                 & GetKyuShikuchoson(dcChosahyo).PadLeft(2, "0"c) _
                 & GetNogyoShuraku(dcChosahyo).PadLeft(3, "0"c) _
                 & GetChosaku(dcChosahyo).PadLeft(3, "0"c) _
                 & GetKyakutaiNo(dcChosahyo).PadLeft(3, "0"c)
            End If

            Return ret
        End Function

        ''' <summary>
        ''' レコードの枝番を取得する
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetEdaNo(ByVal value As String) As String

            Return "_" & GetDetailNo(value)

        End Function

        ''' <summary>
        ''' レコードの明細番号を取得する
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetDetailNo(ByVal value As String) As String

            Return value.Substring(value.IndexOf("_") + 1)

        End Function

        ''' <summary>
        ''' レコードの項目番号を取得する
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetItemNo(ByVal value As String) As String

            Return value.Substring(0, value.IndexOf("_"))

        End Function

        ''' <summary>
        ''' 電子調査票の可変項目の歯抜け行を上に詰める
        ''' </summary>
        ''' <param name="itemAddress">処理対象のシート内の項目の項番(Q07010101等)</param>
        ''' <param name="chosahyoAllItems">電子調査票の全項目</param>
        ''' <param name="sheetName">シート名</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MoveUpMissingRow(ByVal itemAddress As String, ByVal chosahyoAllItems As Dictionary(Of String, DAOChosahyo.調査票項目), Optional sheetName As String = Nothing) As Dictionary(Of String, DAOChosahyo.調査票項目)

            If String.IsNullOrWhiteSpace(itemAddress) OrElse itemAddress.Length < 3 Then
                Return Nothing
            End If

            If Not chosahyoAllItems.Any() Then
                Return Nothing
            End If

            'シート内の可変項目を取得する
            Dim query = chosahyoAllItems.Where(Function(row) row.Key.StartsWith(itemAddress.Substring(0, 3)) AndAlso
                                                                 row.Key.IndexOf("_") = 9).ToList()

            If sheetName <> Nothing Then
                query = chosahyoAllItems.Where(Function(row) row.Key.StartsWith(itemAddress.Substring(0, 3)) AndAlso
                                                                 row.Value.シート名 = sheetName AndAlso
                                                                 row.Key.IndexOf("_") = 9).ToList()
            End If


            'シート内の各項目を明細番号ごとにグループ化したリストを取得する
            Dim getChosahyoGroupListByDetailNo = Function()

                                                     '可変テーブル内の明細番号の一覧をソートして取得する
                                                     'ソートの補足：「_1」⇒「00001」、「_01」⇒「_000001」、「_11111」⇒「11111」と5桁の数値にして昇順ソートする
                                                     Dim query1 = query.Select(Function(row) GetEdaNo(row.Key)).Distinct().OrderBy(Function(row) String.Format("{0:D5}", CInt(row.Substring(1)))).ToList()

                                                     Dim list As New List(Of List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目)))

                                                     '明細番号ごとにグルーピングする
                                                     query1.ForEach(Sub(row) list.Add(query.Where(Function(row2) row2.Key.EndsWith(row)).ToList()))

                                                     Return list

                                                 End Function

            'シート内の各項目の明細番号を振り直す
            Dim executeReNumberingChosahyo = Function(chosahyoGroupList As List(Of List(Of KeyValuePair(Of String, DAOChosahyo.調査票項目))))

                                                 If Not chosahyoGroupList.Any() Then
                                                     Return Nothing
                                                 End If

                                                 Dim chosahyoDictionary As New Dictionary(Of String, DAOChosahyo.調査票項目)
                                                 Dim i = 1 '採番し直す用の行数

                                                 'シート内の値が設定されている行を上から順に処理する
                                                 For Each detailRow In chosahyoGroupList

                                                     '一行まるまる未入力(空白)だったら処理しない
                                                     If detailRow.All(Function(row) row.Value.値 Is Nothing) Then
                                                         Continue For
                                                     End If

                                                     'REV_020↓子牛生産費かつ【２】対象畜概要１の場合、個体識別番号が空だったら処理しない
                                                     If CommonInfo.Chosakubun = ComConst.調査区分.子牛生産費統計_個別 _
                                                         AndAlso sheetName = "【２】対象畜概要１" _
                                                         AndAlso (detailRow.Where(Function(x) x.Key.StartsWith("Q02030301")).Count() = 0 _
                                                         OrElse detailRow.Where(Function(x) x.Key.StartsWith("Q02030301") AndAlso (String.IsNullOrEmpty(x.Value.値))).Count() > 0) Then
                                                         Continue For
                                                     End If
                                                     'REV_020↑

                                                     '歯抜け行となるか
                                                     If Not GetDetailNo(detailRow.First().Key) = CStr(i) Then

                                                         '一項目ずつ処理をする
                                                         For Each column In detailRow

                                                             '「行位置」のみ値を変更する
                                                             Dim chosahyoItem As New DAOChosahyo.調査票項目 With {.シート名 = column.Value.シート名,
                                                                                                                  .可変範囲 = column.Value.可変範囲,
                                                                                                                  .型区分 = column.Value.型区分,
                                                                                                                  .行位置 = i,
                                                                                                                  .小数点以下桁数 = column.Value.小数点以下桁数,
                                                                                                                  .値 = column.Value.値,
                                                                                                                  .有効桁数 = column.Value.有効桁数,
                                                                                                                  .列位置 = column.Value.列位置}

                                                             chosahyoDictionary.Add(GetItemNo(column.Key) & "_" & CStr(i), chosahyoItem)

                                                         Next
                                                     End If

                                                     i += 1
                                                 Next

                                                 Return chosahyoDictionary

                                             End Function

            '行を上に詰める
            Dim moveUpChosahyo = Sub(reNumberingItems As Dictionary(Of String, DAOChosahyo.調査票項目))

                                     If Not reNumberingItems.Any() Then
                                         Return
                                     End If

                                     '上書き対象クリア
                                     Dim rowNoList = (From item In reNumberingItems Select item.Key.Split(CChar(ComConst.ITEM_NO_DELIMITER))(1)).Distinct

                                     For Each rowNo In rowNoList
                                         Dim query0 = query.Where(Function(row) row.Key Like "*" & ComConst.ITEM_NO_DELIMITER & rowNo).ToList()
                                         For Each kv In query0
                                             kv.Value.値 = Nothing
                                         Next
                                     Next

                                     Dim query1 = query.ToDictionary(Function(row) row.Key, Function(row) row.Value)

                                     '採番し直した項目の数ぶんループする
                                     For Each reNumberingItem In reNumberingItems

                                         '「採番前の項番」と「採番後の項番」がバッティングする場合は上書きする
                                         If query1.ContainsKey(reNumberingItem.Key) Then
                                             chosahyoAllItems(reNumberingItem.Key) = reNumberingItem.Value
                                         Else
                                             'バッティングしない場合は新規に追加する
                                             chosahyoAllItems.Add(reNumberingItem.Key, reNumberingItem.Value)
                                         End If
                                     Next

                                 End Sub

            '上に詰めたことで不要となった行を削除する
            Dim removeUnnecessaryChosahyo = Sub(reNumberingItems As Dictionary(Of String, DAOChosahyo.調査票項目))

                                                If Not reNumberingItems.Any() Then
                                                    Return
                                                End If

                                                '上詰めした最終行を取得する
                                                Dim lastDetailNo = CInt(GetDetailNo(reNumberingItems.Last().Key))

                                                '上詰めしたことで不要になった項目を取得する
                                                Dim query1 = query.Where(Function(row) lastDetailNo < CInt(GetDetailNo(row.Key))).ToList()

                                                '不要な項目をすべて削除する
                                                query1.ForEach(Sub(row) chosahyoAllItems(row.Key).値 = Nothing)

                                            End Sub

            'シート内の各項目を明細番号ごとにグループ化したリストを取得する
            Dim oldChosahyoItems = getChosahyoGroupListByDetailNo()

            'シート内の各項目の明細番号を振り直す
            Dim newChosahyoItems = executeReNumberingChosahyo(oldChosahyoItems)

            '行を上に詰める
            moveUpChosahyo(newChosahyoItems)

            '上に詰めたことで不要となった行を削除する
            removeUnnecessaryChosahyo(newChosahyoItems)

            Return chosahyoAllItems

        End Function

    End Class

    ''' <summary>
    ''' 個別結果表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkahyo

        ''' <summary>
        ''' 調査年コンボボックス設定
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="db"></param>
        ''' <param name="koutei"></param>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <param name="kessokuchiHokana"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetChosaNenComboBox(chosaNen As ComboBox, db As DBAccess, koutei As String, kyoku As String, jimusho As String, kyoten As String, Optional kessokuchiHokana As String = ComConst.欠測値補完.無)
            Dim dt As DataTable = Nothing
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOKobetsuKekkahyo.GetChosaNen(db, kyoku, jimusho, kyoten, kessokuchiHokana)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOKobetsuKekkahyo.GetChosaNen(db, kyoku, Nothing, Nothing, kessokuchiHokana)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOKobetsuKekkahyo.GetChosaNen(db, Nothing, Nothing, Nothing, kessokuchiHokana)
            End Select

            chosaNen.ValueMember = "調査年"
            chosaNen.DisplayMember = "調査年"
            chosaNen.DataSource = dt
        End Sub

        ''' <summary>
        ''' 調査年コンボボックス設定（農業経営体対応）
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="db"></param>
        ''' <param name="koutei"></param>
        ''' <param name="chosakubun"></param>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <param name="kessokuchiHokana"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetChosaNenComboBox(chosaNen As ComboBox, db As DBAccess, koutei As String, chosakubun As String, kyoku As String, jimusho As String, kyoten As String, kessokuchiHokana As String)
            If Not chosakubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                SetChosaNenComboBox(chosaNen, db, koutei, kyoku, jimusho, kyoten, kessokuchiHokana)
            Else
                Dim dtMerge As New DataTable
                Dim dtWork As DataTable = Nothing
                For Each arr As String In {ComConst.営農経営体区分.個人経営体, ComConst.営農経営体区分.法人経営体}
                    Select Case koutei
                        Case CommonInfo.KouteiKubun.Code.Center
                            dtWork = DAOKobetsuKekkahyo.GetChosaNen(db, arr, kyoku, jimusho, kyoten, kessokuchiHokana)
                        Case CommonInfo.KouteiKubun.Code.Kyoku
                            dtWork = DAOKobetsuKekkahyo.GetChosaNen(db, arr, kyoku, Nothing, Nothing, kessokuchiHokana)
                        Case CommonInfo.KouteiKubun.Code.Honsyo
                            dtWork = DAOKobetsuKekkahyo.GetChosaNen(db, arr, Nothing, Nothing, Nothing, kessokuchiHokana)
                    End Select
                    dtMerge.Merge(dtWork)
                Next

                Dim dv As DataView = New DataView(dtMerge)
                dv.Sort = "調査年 DESC"

                chosaNen.ValueMember = "調査年"
                chosaNen.DisplayMember = "調査年"
                chosaNen.DataSource = dv.ToTable(True)
            End If

        End Sub

        ''' <summary>
        ''' 調査年コンボボックス設定（受信）
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="db"></param>
        ''' <param name="koutei"></param>
        ''' <param name="upLow"></param>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetChosaNenJushinComboBox(chosaNen As ComboBox, db As DBAccess, koutei As String, upLow As String, kyoku As String, jimusho As String, kyoten As String)
            Dim dt As DataTable = Nothing
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOKobetsuKekkahyo.GetChosaNenJushin(db, upLow, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOKobetsuKekkahyo.GetChosaNenJushin(db, upLow, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOKobetsuKekkahyo.GetChosaNenJushin(db, upLow, Nothing, Nothing, Nothing)
            End Select

            chosaNen.ValueMember = "調査年"
            chosaNen.DisplayMember = "調査年"
            chosaNen.DataSource = dt
        End Sub

        ''' <summary>
        ''' 個別結果表項目取得
        ''' </summary>
        ''' <param name="dtItem"></param>
        ''' <param name="dc"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetItem(dtItem As DataTable, dc As Dictionary(Of String, DataTable), Optional nousanFlg As Boolean = False) As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
            Dim ret As New Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

            Dim lstItem As List(Of DataRow) = dtItem.AsEnumerable().ToList()

            For Each kv As KeyValuePair(Of String, DataTable) In dc
                Dim dt As DataTable = kv.Value
                Dim lstDr As List(Of DataRow) = dt.AsEnumerable().ToList()
                Dim colArr(dt.Columns.Count - 1) As DataColumn
                dt.Columns.CopyTo(colArr, 0)

                For Each row As DataRow In lstDr
                    For Each col As DataColumn In colArr
                        Dim query = (From dr In lstItem Where dr("項目番号").ToString = col.ColumnName Select dr).Take(1).ToArray
                        If query.Any() Then
                            Dim item As New DAOKobetsuKekkahyo.個別結果表項目
                            With item
                                .シート名 = query(0)("シート名").ToString
                                .行位置 = Integer.Parse(query(0)("行位置").ToString)
                                .列位置 = Integer.Parse(query(0)("列位置").ToString)
                                .表示単位 = query(0)("表示単位").ToString
                                .型区分 = query(0)("型区分").ToString
                                'REV_012 START---------------
                                '農産物生産費の場合、制度受取金・積立金等項目名を設定する
                                If nousanFlg Then
                                    If query(0)("出力項目名").ToString <> "" Then
                                        .値 = query(0)("出力項目名").ToString
                                    Else
                                        .値 = row(col.ColumnName).ToString
                                    End If
                                Else
                                    .値 = row(col.ColumnName).ToString
                                End If
                                'REV_012 END---------------
                            End With
                            ret.Add(col.ColumnName, item)
                        End If
                    Next
                Next
            Next

            Return ret
        End Function

        ''' <summary>
        ''' 個別結果表シートデータ設定
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dc As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess)
            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object

                        Dim page As Excel.PageSetup = xlSheet.PageSetup
                        Try
                            rng = xlSheet.Range(page.PrintArea)
                        Finally
                            ReleaseComObject(comObject, page)
                        End Try

                        arrData = DirectCast(rng.Formula, Object(,))

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 Select dr

                        For Each kv As KeyValuePair(Of String, DAOKobetsuKekkahyo.個別結果表項目) In query
                            arrData(kv.Value.行位置, kv.Value.列位置) = GetformattedValue(kv.Value)
                        Next

                        rng.Value = arrData
                        rng.Value = rng.Formula
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' フォーマット済値取得
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetformattedValue(item As DAOKobetsuKekkahyo.個別結果表項目) As String
            Dim ret As String

            Dim format As String = String.Empty

            If Not String.IsNullOrEmpty(item.表示単位) Then
                Dim unit As Decimal = Decimal.Parse(item.表示単位)
                If ComConst.個別結果表作成論理.表示単位.リスト.ContainsKey(unit) Then
                    format = ComConst.個別結果表作成論理.表示単位.リスト(unit)
                End If
            End If

            Dim val As Decimal
            If Not format.Equals(String.Empty) AndAlso Decimal.TryParse(item.値, val) Then
                ret = val.ToString(format)
            Else
                ret = item.値
            End If

            Return ret
        End Function

        ''' <summary>
        ''' 集計用テーブル付加名称取得
        ''' </summary>
        ''' <param name="chosakubun"></param>
        ''' <param name="kessokuchiHokana"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSyukeiTableAddName(chosakubun As String, kessokuchiHokana As String) As String
            Dim ret As String = String.Empty

            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo _
                And chosakubun = ComConst.調査区分.営農類型別経営統計_個人 _
                And kessokuchiHokana.Equals(ComConst.欠測値補完.有) Then
                ret = ComConst.個別結果表.集計用テーブル付加名称
            End If

            Return ret
        End Function

        ''' <summary>
        ''' 値を取得する
        ''' </summary>
        ''' <param name="value">値</param>
        ''' <param name="type">項目の型</param>
        ''' <param name="isCnvDec">値を数値変換するかどうか</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetData(ByVal value As Object, ByVal type As String, ByVal isCnvDec As Boolean) As String
            Dim ret As String

            If value Is Nothing OrElse String.IsNullOrEmpty(value.ToString) Then
                ret = Nothing
            Else
                '数値項目で数値変換が指定された場合
                If type = ComConst.型区分.数値型 AndAlso isCnvDec Then
                    Dim cnvDecimal As Decimal
                    If Decimal.TryParse(CStr(value), cnvDecimal) Then
                        ret = cnvDecimal.ToString
                    Else
                        ret = Nothing
                    End If
                Else
                    ret = value.ToString
                End If
            End If

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 個別結果検討表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkaKentohyo

        ''' <summary>
        ''' 個別結果検討表項目クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 個別結果検討表項目
            Public シート名 As String                  'シート名
            Public 行位置 As Integer                   '行位置
            Public 列位置 As Integer                   '列位置
            Public 値 As String                        '値
            Public 型区分 As String                    '型区分
            Public 有効桁数 As Integer                 '有効桁数
            Public 小数点以下桁数 As Integer           '小数点以下桁数
            Public 表示単位 As String                  '表示単位
        End Class

        ''' <summary>
        ''' 個別結果検討表シートデータ設定
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dc As Dictionary(Of String, 個別結果検討表項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess)
            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object

                        Dim page As Excel.PageSetup = xlSheet.PageSetup
                        Try
                            rng = xlSheet.Range(page.PrintArea)
                        Finally
                            ReleaseComObject(comObject, page)
                        End Try

                        arrData = DirectCast(rng.Formula, Object(,))

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 Select dr

                        For Each kv As KeyValuePair(Of String, 個別結果検討表項目) In query
                            Dim format As String = String.Empty

                            If Not String.IsNullOrEmpty(kv.Value.表示単位) Then
                                Dim unit As Decimal = Decimal.Parse(kv.Value.表示単位)
                                If ComConst.個別結果検討表作成論理.表示単位.リスト.ContainsKey(unit) Then
                                    format = ComConst.個別結果検討表作成論理.表示単位.リスト(unit)
                                End If
                            End If

                            Dim val As Decimal
                            If Not format.Equals(String.Empty) AndAlso Decimal.TryParse(kv.Value.値, val) Then
                                arrData(kv.Value.行位置, kv.Value.列位置) = val.ToString(format)
                            Else
                                arrData(kv.Value.行位置, kv.Value.列位置) = kv.Value.値
                            End If
                        Next

                        rng.Value = arrData
                        rng.Value = rng.Formula
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
        End Sub
    End Class

    ''' <summary>
    ''' 調査票審査論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ChosahyoShinsaRonri

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.調査票審査論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.調査票審査論理.出力用ファイル名称.Col.First & ComConst.調査票審査論理.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.調査票審査論理.出力用ファイル名称.Col.First & ComConst.調査票審査論理.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.調査票審査論理.出力用ファイル名称.Col.First & ComConst.調査票審査論理.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.調査票審査論理.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.調査票審査論理.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.調査票審査論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.調査票審査論理.出力用ファイル名称.Col.First & ComConst.調査票審査論理.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.調査票審査論理.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.調査票審査論理.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.調査票審査論理.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <param name="errType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), ByVal errType As ComConst.エラーチェック種別.enm) As Boolean
            Dim ret As Boolean = True

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「チェック項目名」、「内容」、「エラーとなる条件」、「区分」は全て入力してください。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーサイン」が半角英数字４桁で入力されておりません。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーサイン」がファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「繰り返し」に「○」以外が入力されております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「区分」に「Z」または「W」以外が入力されております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーとなる条件」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーとなる条件」の条件式が不正です。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「繰り返し」に「○」が入力されておりますが、「エラーとなる条件」に繰り返し項目が1つも存在しません。"
            }

            Dim row As Integer = 0
            Dim cnt As Integer = 0

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                Dim dtItemMst As DataTable = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, 調査年_引数用)
                '審査情報クラス生成
                Dim shinsa As Shinsa = New Shinsa(db,
                                                  CommonInfo.Chosakubun,
                                                  ComConst.審査論理データ種別.調査票,
                                                  If(errType = ComConst.エラーチェック種別.enm.基本, ComConst.審査論理種別.基本チェック, ComConst.審査論理種別.追加チェック),
                                                  dtItemMst)

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    '1）チェック項目名、内容、エラーとなる条件、区分については、全て入力されているか。
                    If dc("チェック項目名").ToString.Equals(String.Empty) _
                        OrElse dc("エラー内容").ToString.Equals(String.Empty) _
                        OrElse dc("エラーとなる条件").ToString.Equals(String.Empty) _
                        OrElse dc("エラー区分").ToString.Equals(String.Empty) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）エラーサインが半角英数字４桁であるか。
                    If Not Regex.IsMatch(dc("エラーサイン").ToString, "^[0-9a-zA-Z]{4}$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '3）エラーサインがファイル上で重複していないか。
                    Dim query = From dct In lstDc Where dct("エラーサイン").ToString = dc("エラーサイン").ToString
                    If query.Count > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '4）繰り返しに「○」以外の入力があるか。
                    If Not dc("繰り返し").ToString.Equals(String.Empty) Then
                        If Not dc("繰り返し").ToString.Equals("○") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '5）区分が「Z」または「W」でがあるか。
                    If Not dc("エラー区分").ToString.Equals(String.Empty) Then
                        If Not (dc("エラー区分").ToString.Equals("Z") Or dc("エラー区分").ToString.Equals("W")) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '6）エラーとなる条件の入力が正しいかどうか。
                    'a)入力されている項番が存在するか
                    If shinsa.CheckExistItemNo(dc("エラーサイン").ToString, dc("エラーとなる条件").ToString) Then
                        'b)審査を実行できる条件となっているか。（SQLエラーとならない事をチェックする。）
                        '「エラーとなる条件」のチェック
                        If Not shinsa.CheckExecutableSQL(dc("エラーサイン").ToString, dc("エラーとなる条件").ToString, dc("繰り返し").ToString) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                        'c)「繰り返し」に「○」がチェックされているが、エラーとなる条件に繰り返し項目が1つも存在しない。
                        '「エラーとなる条件」及び「繰り返し」のチェック
                        If Not shinsa.CheckRepeat(dc("エラーサイン").ToString, dc("エラーとなる条件").ToString, dc("繰り返し").ToString) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    Else
                        cnt = cnt + 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 個別結果表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkahyoSakuseiRonri

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.個別結果表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.個別結果表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.個別結果表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.個別結果表作成論理.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表作成論理.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.個別結果表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.個別結果表作成論理.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.個別結果表作成論理.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表作成論理.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ' REV_011↓
        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <param name="versionKbn"></param>
        ''' <returns></returns>
        'Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String)) As Boolean
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), versionKbn As String) As Boolean
            ' REV_011↑
            Dim ret As Boolean = True
            Dim dtChoItemMst As DataTable
            Dim dtKobetsuItemMst As DataTable

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　項番（{2}）の「出力項目名」、「優先順位」、「計算式」は全て入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」がファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「優先順位」は半角数値0.1～99.9で入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」の条件式が不正です。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「表示単位」は半角数値で9、9.9、9.99、9.999のいずれかで入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「再計算」に「○」以外が入力されております。"
            }

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, 調査年_引数用)
                '個別結果表項目マスタ取得(裏項番含める)
                ' REV_011↓
                'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn, True)
                ' REV_011↑

                Dim dtCreateRonri As New DataTable
                dtCreateRonri.Columns.Add("計算式", GetType(String))
                For Each dc As Dictionary(Of String, String) In lstDc
                    If Not String.IsNullOrEmpty(dc("計算式").ToString) Then
                        dtCreateRonri.Rows.Add(dc("計算式").ToString)
                    End If
                Next

                '個別結果表・個別結果検討表作成クラス
                Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                                 CommonInfo.Chosakubun,
                                                                 CommonInfo.Chosakubun,
                                                                 CreateKobetsu.enmCreateType.個別結果表作成,
                                                                 dtChoItemMst,
                                                                 dtKobetsuItemMst,
                                                                 Nothing,
                                                                 dtCreateRonri)

                Dim row As Integer = 0
                Dim cnt As Integer = 0
                '項番が存在するかのチェック用
                Dim isExistItemNo As Boolean

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    isExistItemNo = False

                    '1）出力項目名、優先順位、計算式、については、全て入力されているか。
                    If dc("項目名").ToString.Equals(String.Empty) _
                        OrElse dc("優先順位").ToString.Equals(String.Empty) _
                        OrElse dc("計算式").ToString.Equals(String.Empty) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）項番が存在するか。
                    Dim query1 = From dr In dtKobetsuItemMst Where dr("項目番号").ToString = dc("項目番号").ToString Select dr
                    If Not query1.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        isExistItemNo = True
                    End If

                    '3）項番がファイル上で重複していないか。
                    Dim query2 = From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString
                    If query2.Count > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '4）優先順位が半角数値で0.1から99.9までの入力であること。
                    If Not Regex.IsMatch(dc("優先順位").ToString, "^[0-9]{1,2}(\.[0-9]{1})?$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("優先順位"), val) Then
                            If Not (val >= 0.1 And val <= 99.9) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If

                    '項番が存在する場合
                    If isExistItemNo Then
                        '5）計算式の入力が正しいかどうか。
                        'a)計算式として設定されている項番が存在するか
                        If kobetsu.CheckExistItemNo(dc("項目番号").ToString, dc("計算式").ToString) Then
                            'b)個別結果表作成を実行できる計算式となっているか。（SQLエラーとならない事をチェックする。）
                            If Not kobetsu.CheckExecutableSQL(dc("項目番号").ToString, dc("計算式").ToString) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        Else
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '6）表示単位が半角数値で9、9.9、9.99、9.999のいずれかの入力であること。
                    If Not dc("表示単位").ToString.Equals(String.Empty) Then
                        If Not Regex.IsMatch(dc("表示単位").ToString, "^[0-9]{1}(\.[0-9]{1,3})?$") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        Else
                            Dim val As Decimal
                            If Decimal.TryParse(dc("表示単位"), val) Then
                                If Not ComConst.個別結果表作成論理.表示単位.リスト.ContainsKey(val) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If

                    '7）再計算に「○」以外の入力があるか。
                    If Not dc("再計算区分").ToString.Equals(String.Empty) Then
                        If Not dc("再計算区分").ToString.Equals("○") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 個別結果表作成論理（営農個人）クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkahyoSakuseiRonriEinouKobetsu

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Col.First & ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表作成論理_営農個人.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ' REV_011↓
        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <param name="versionKbn"></param>
        ''' <returns></returns>
        'Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String)) As Boolean
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), versionKbn As String) As Boolean
            ' REV_011↑
            Dim ret As Boolean = True
            Dim dtChoItemMst As DataTable
            Dim dtKobetsuItemMst As DataTable

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　項番（{2}）の「出力項目名」、「優先順位」、「計算式」は全て入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「貸借対照表区分」は半角数値0～3で入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」と「貸借対照表区分」の組み合わせがファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「貸借対照表区分：0」の「項番」が重複しております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「貸借対照表区分：0～3」の「項番」が全て存在しておりません。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「優先順位」は半角数値0.1～99.9で入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」の条件式が不正です。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「表示単位」は半角数値で9、9.9、9.99、9.999のいずれかで入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「再計算」に「○」以外が入力されております。"
            }

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, 調査年_引数用)
                '個別結果表項目マスタ取得(裏項番含める)
                ' REV_011↓
                'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn, True)
                ' REV_011↑

                Dim dtCreateRonri As New DataTable
                dtCreateRonri.Columns.Add("計算式", GetType(String))
                For Each dc As Dictionary(Of String, String) In lstDc
                    If Not String.IsNullOrEmpty(dc("計算式").ToString) Then
                        dtCreateRonri.Rows.Add(dc("計算式").ToString)
                    End If
                Next

                '個別結果表・個別結果検討表作成クラス
                Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                                 CommonInfo.Chosakubun,
                                                                 CommonInfo.Chosakubun,
                                                                 CreateKobetsu.enmCreateType.個別結果表作成,
                                                                 dtChoItemMst,
                                                                 dtKobetsuItemMst,
                                                                 Nothing,
                                                                 dtCreateRonri)

                Dim row As Integer = 0
                Dim cnt As Integer = 0
                '項番が存在するかのチェック用
                Dim isExistItemNo As Boolean

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    isExistItemNo = False

                    '1）出力項目名、優先順位、計算式、については、全て入力されているか。
                    If dc("項目名").ToString.Equals(String.Empty) _
                        OrElse dc("優先順位").ToString.Equals(String.Empty) _
                        OrElse dc("計算式").ToString.Equals(String.Empty) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）項番が存在するか。
                    Dim query1 = From dr In dtKobetsuItemMst Where dr("項目番号").ToString = dc("項目番号").ToString Select dr
                    If Not query1.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        isExistItemNo = True
                    End If

                    '3）貸借対照表区分が0～3のいづれかであるか。
                    If Not dc("貸借対照表区分").ToString.Equals(String.Empty) Then
                        If Not Regex.IsMatch(dc("貸借対照表区分").ToString, "^[0-9]{1}$") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        Else
                            If Not ComConst.個別結果表作成論理_営農個人.貸借対照表区分.リスト.ContainsKey(dc("貸借対照表区分")) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If

                    '4）項番と貸借対照表区分の組み合わせがファイル上で重複していないか。
                    Dim recCount As Integer
                    recCount = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString And dct("貸借対照表区分").ToString = dc("貸借対照表区分").ToString).Count
                    If recCount > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '5）貸借対照表区分0の項目は重複していないか。
                    If dc("貸借対照表区分").Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._0) Then
                        recCount = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString And Not dct("貸借対照表区分").ToString.Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._0)).Count
                        If recCount > 0 Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '6）貸借対照表区分1～3の項目は全て存在しているか。
                    If dc("貸借対照表区分").Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._1) _
                        OrElse dc("貸借対照表区分").Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._2) _
                        OrElse dc("貸借対照表区分").Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._3) Then

                        Dim recCount1 As Integer = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString And dct("貸借対照表区分").ToString.Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._1)).Count
                        Dim recCount2 As Integer = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString And dct("貸借対照表区分").ToString.Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._2)).Count
                        Dim recCount3 As Integer = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString And dct("貸借対照表区分").ToString.Equals(ComConst.個別結果表作成論理_営農個人.貸借対照表区分._3)).Count

                        If recCount1 <> 1 OrElse recCount2 <> 1 OrElse recCount3 <> 1 Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '7）優先順位が半角数値で0.1から99.9までの入力であること。
                    If Not Regex.IsMatch(dc("優先順位").ToString, "^[0-9]{1,2}(\.[0-9]{1})?$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("優先順位"), val) Then
                            If Not (val >= 0.1 And val <= 99.9) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If

                    '項番が存在する場合
                    If isExistItemNo Then
                        '8）計算式の入力が正しいかどうか。
                        'a)計算式として設定されている項番が存在するか
                        If kobetsu.CheckExistItemNo(dc("項目番号").ToString, dc("計算式").ToString) Then
                            'b)個別結果表作成を実行できる計算式となっているか。（SQLエラーとならない事をチェックする。）
                            If Not kobetsu.CheckExecutableSQL(dc("項目番号").ToString, dc("計算式").ToString) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        Else
                            cnt = cnt + 1
                            details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '9）表示単位が半角数値で9、9.9、9.99、9.999のいずれかの入力であること。
                    If Not dc("表示単位").ToString.Equals(String.Empty) Then
                        If Not Regex.IsMatch(dc("表示単位").ToString, "^[0-9]{1}(\.[0-9]{1,3})?$") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        Else
                            Dim val As Decimal
                            If Decimal.TryParse(dc("表示単位"), val) Then
                                If Not ComConst.個別結果表作成論理_営農個人.表示単位.リスト.ContainsKey(val) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If

                    '10）再計算に「○」以外の入力があるか。
                    If Not dc("再計算区分").ToString.Equals(String.Empty) Then
                        If Not dc("再計算区分").ToString.Equals("○") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(11), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 個別結果検討表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkaKentohyoSakuseiRonri

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果検討表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.個別結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果検討表作成論理.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.個別結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果検討表作成論理.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.個別結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果検討表作成論理.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.個別結果検討表作成論理.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果検討表作成論理.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果検討表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.個別結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.個別結果検討表作成論理.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.個別結果検討表作成論理.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.個別結果検討表作成論理.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果検討表作成論理.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ' REV_011↓
        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <param name="versionKbn"></param>
        ''' <returns></returns>
        'Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String)) As Boolean
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), versionKbn As String) As Boolean
            ' REV_011↑
            Dim ret As Boolean = True
            Dim dtKobetsuItemMst As DataTable
            Dim dtKobetsuKekkaItemMst As DataTable

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　項番（{2}）の「出力項目名」、「優先順位」、「計算式」は全て入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」がファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「優先順位」は半角数値0.1～99.9で入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」の条件式が不正です。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「表示単位」は半角数値で9、9.9、9.99、9.999のいずれかで入力してください。"
            }

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表項目マスタ取得(裏項番含める)
                ' REV_011↓
                'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn, True)
                ' REV_011↑
                '個別結果検討表項目マスタ取得
                ' REV_011↓
                'dtKobetsuKekkaItemMst = DAOOther.GetKobetsuKekkaKentohyoItemMaster(db, CommonInfo.Chosakubun)
                dtKobetsuKekkaItemMst = DAOOther.GetKobetsuKekkaKentohyoItemMaster(db, CommonInfo.Chosakubun, versionKbn)
                ' REV_011↑

                '個別結果表・個別結果検討表作成クラス
                Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                                 CommonInfo.Chosakubun,
                                                                 CommonInfo.Chosakubun,
                                                                 CreateKobetsu.enmCreateType.個別結果検討表作成,
                                                                 Nothing,
                                                                 dtKobetsuItemMst,
                                                                 dtKobetsuKekkaItemMst,
                                                                 Nothing)

                Dim row As Integer = 0
                Dim cnt As Integer = 0
                '項番が存在するかのチェック用
                Dim isExistItemNo As Boolean

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    isExistItemNo = False

                    '1）出力項目名、優先順位、計算式、については、全て入力されているか。
                    If dc("項目名").ToString.Equals(String.Empty) _
                        OrElse dc("優先順位").ToString.Equals(String.Empty) _
                        OrElse dc("計算式").ToString.Equals(String.Empty) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）項番が存在するか。
                    Dim query1 = From dr In dtKobetsuKekkaItemMst Where dr("項目番号").ToString = dc("項目番号").ToString Select dr
                    If Not query1.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        isExistItemNo = True
                    End If

                    '3）項番がファイル上で重複していないか。
                    Dim query2 = From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString
                    If query2.Count > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '4）優先順位が半角数値で0.1から99.9までの入力であること。
                    If Not Regex.IsMatch(dc("優先順位").ToString, "^[0-9]{1,2}(\.[0-9]{1})?$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("優先順位"), val) Then
                            If Not (val >= 0.1 And val <= 99.9) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If

                    '項番が存在する場合
                    If isExistItemNo Then
                        '5）計算式の入力が正しいかどうか。
                        'a)計算式として設定されている項番が存在するか
                        If kobetsu.CheckExistItemNo(dc("項目番号").ToString, dc("計算式").ToString) Then
                            'b)個別結果検討表作成を実行できる計算式となっているか。（SQLエラーとならない事をチェックする。）
                            If Not kobetsu.CheckExecutableSQL(dc("項目番号").ToString, dc("計算式").ToString) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        Else
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '6）表示単位が半角数値で9、9.9、9.99、9.999のいずれかの入力であること。
                    If Not dc("表示単位").ToString.Equals(String.Empty) Then
                        If Not Regex.IsMatch(dc("表示単位").ToString, "^[0-9]{1}(\.[0-9]{1,3})?$") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        Else
                            Dim val As Decimal
                            If Decimal.TryParse(dc("表示単位"), val) Then
                                If Not ComConst.個別結果検討表作成論理.表示単位.リスト.ContainsKey(val) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If
                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 個別結果表審査論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkahyoShinsaRonri

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表審査論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.個別結果表審査論理.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.個別結果表審査論理.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.個別結果表審査論理.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.個別結果表審査論理.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表審査論理.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表審査論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.個別結果表審査論理.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.個別結果表審査論理.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.個別結果表審査論理.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表審査論理.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ' REV_011↓
        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <param name="errType"></param>
        ''' <param name="versionKbn"></param>
        ''' <returns></returns>
        'Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), ByVal errType As ComConst.エラーチェック種別.enm) As Boolean
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), ByVal errType As ComConst.エラーチェック種別.enm, versionKbn As String) As Boolean
            ' REV_011↑
            Dim ret As Boolean = True

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「チェック項目名」、「内容」、「エラーとなる条件」、「区分」は全て入力してください。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーサイン」が半角英数字４桁で入力されておりません。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーサイン」がファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「区分」に「Z」または「W」以外が入力されております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーとなる条件」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　エラーサイン（{2}）の「エラーとなる条件」の条件式が不正です。"
            }

            Dim row As Integer = 0
            Dim cnt As Integer = 0

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表項目マスタ取得
                ' REV_011↓
                'Dim dtItemMst As DataTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                Dim dtItemMst As DataTable = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn)
                ' REV_011↑
                '審査情報クラス生成
                ' REV_011↓
                'Dim shinsa As Shinsa = New Shinsa(db,
                '                                  CommonInfo.Chosakubun,
                '                                  ComConst.審査論理データ種別.個別結果表,
                '                                  If(errType = ComConst.エラーチェック種別.enm.基本, ComConst.審査論理種別.基本チェック, ComConst.審査論理種別.追加チェック),
                '                                  dtItemMst)
                Dim shinsa As Shinsa = New Shinsa(db,
                                                  CommonInfo.Chosakubun,
                                                  ComConst.審査論理データ種別.個別結果表,
                                                  If(errType = ComConst.エラーチェック種別.enm.基本, ComConst.審査論理種別.基本チェック, ComConst.審査論理種別.追加チェック),
                                                  dtItemMst,
                                                  versionKbn)
                ' REV_011↑

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    '1）チェック項目名、内容、エラーとなる条件、区分については、全て入力されているか。
                    If dc("チェック項目名").ToString.Equals(String.Empty) _
                        OrElse dc("エラー内容").ToString.Equals(String.Empty) _
                        OrElse dc("エラーとなる条件").ToString.Equals(String.Empty) _
                        OrElse dc("エラー区分").ToString.Equals(String.Empty) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）エラーサインが半角英数字４桁であるか。
                    If Not Regex.IsMatch(dc("エラーサイン").ToString, "^[0-9a-zA-Z]{4}$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '3）エラーサインがファイル上で重複していないか。
                    Dim query = From dct In lstDc Where dct("エラーサイン").ToString = dc("エラーサイン").ToString
                    If query.Count > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '4）区分が「Z」または「W」でがあるか。
                    If Not dc("エラー区分").ToString.Equals(String.Empty) Then
                        If Not (dc("エラー区分").ToString.Equals("Z") Or dc("エラー区分").ToString.Equals("W")) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '5）エラーとなる条件の入力が正しいかどうか。
                    'a)入力されている項番が存在するか
                    If shinsa.CheckExistItemNo(dc("エラーサイン").ToString, dc("エラーとなる条件").ToString) Then
                        'b)審査を実行できる条件となっているか。（SQLエラーとならない事をチェックする。）
                        '「エラーとなる条件」のチェック
                        If Not shinsa.CheckExecutableSQL(dc("エラーサイン").ToString, dc("エラーとなる条件").ToString) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    Else
                        cnt = cnt + 1
                        details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("エラーサイン").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 個別結果表審査論理範囲クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KobetsuKekkahyoShinsaRonriRange

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表審査論理範囲.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.個別結果表審査論理範囲.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.個別結果表審査論理範囲.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.個別結果表審査論理範囲.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.個別結果表審査論理範囲.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表審査論理範囲.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            dc("連番") = i.ToString
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.個別結果表審査論理範囲.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.個別結果表審査論理範囲.出力用ファイル名称.Col.First & ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.個別結果表審査論理範囲.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.個別結果表審査論理範囲.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.個別結果表審査論理範囲.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula
                Finally
                    ReleaseComObject(comObject, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try
        End Sub

        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), versionKbn As String) As Boolean
            Dim ret As Boolean = True

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim val1 As Long
            Dim val2 As Decimal
            Dim val3 As Decimal

            '---REV002 MOD START
            'Dim msg As String() = {"" _
            '                     , "{0}件目：{1}行目　「項目番号①」、「項目番号②」、「値③」、「データ範囲下限」、「データ範囲上限」は全て入力してください。" _
            '                     , "{0}件目：{1}行目　「項目番号①」に存在しない項番が入力されております。" _
            '                     , "{0}件目：{1}行目　「項目番号②」に存在しない項番が入力されております。" _
            '                     , "{0}件目：{1}行目　「値③」は正の整数5桁までで入力してください。" _
            '                     , "{0}件目：{1}行目　「データ範囲下限」は正の整数15桁、小数2桁までで入力してください。" _
            '                     , "{0}件目：{1}行目　「データ範囲上限」は正の整数15桁、小数2桁までで入力してください。" _
            '                     , "{0}件目：{1}行目　「データ範囲上限」は「データ範囲下限」より大きな値で入力してください。" _
            '}
            Dim msg As String() = {"" _
                     , "{0}件目：{1}行目　「項目番号①」、「項目番号②」、「値③」、「データ範囲下限」、「データ範囲上限」は全て入力してください。" _
                     , "{0}件目：{1}行目　「項目番号①」に存在しない項番が入力されております。" _
                     , "{0}件目：{1}行目　「項目番号②」に存在しない項番が入力されております。" _
                     , "{0}件目：{1}行目　「値③」は正の整数5桁までで入力してください。" _
                     , "{0}件目：{1}行目　「データ範囲下限」は整数15桁、小数2桁までで入力してください。" _
                     , "{0}件目：{1}行目　「データ範囲上限」は整数15桁、小数2桁までで入力してください。" _
                     , "{0}件目：{1}行目　「データ範囲上限」は「データ範囲下限」より大きな値で入力してください。"
            }
            '---REV002 MOD END

            '個別結果表項目マスタ取得
            Dim dtItem As DataTable
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                ' REV_011↓
                'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, versionKbn)
                ' REV_011↑
            End Using

            Dim cnt As Integer = 0

            For Each dc As Dictionary(Of String, String) In lstDc
                '1）項目番号①、項目番号②、値③、データ範囲下限、データ範囲上限については、全て入力されているか。
                If dc("項目番号１").ToString.Equals(String.Empty) _
                    OrElse dc("項目番号２").ToString.Equals(String.Empty) _
                    OrElse dc("値").ToString.Equals(String.Empty) _
                    OrElse dc("下限").ToString.Equals(String.Empty) _
                    OrElse dc("上限").ToString.Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                    ret = False
                    If cnt = max Then Return ret
                End If

                '2）項目番号①が存在するか。
                If Not dc("項目番号１").ToString.Equals(String.Empty) Then
                    'REV_022↓
                    'Dim query = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号１").ToString Select dr
                    Dim itemNo = dc("項目番号１").ToString
                    If itemNo.StartsWith("前") Then
                        itemNo = itemNo.Substring(1)
                    End If
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = itemNo Select dr
                    'REV_022↑
                    If Not query.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '3）項目番号②が存在するか。
                If Not dc("項目番号２").ToString.Equals(String.Empty) Then
                    'REV_022↓
                    'Dim query = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号２").ToString Select dr
                    Dim itemNo = dc("項目番号２").ToString
                    If itemNo.StartsWith("前") Then
                        itemNo = itemNo.Substring(1)
                    End If
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = itemNo Select dr
                    'REV_022↑
                    If Not query.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '4）値③が5桁までの正数であるか。
                If Not dc("値").ToString.Equals(String.Empty) Then
                    If Long.TryParse(dc("値"), val1) Then
                        If Not (val1 > 0 And val1 < 10 ^ 5) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    Else
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '5）データ範囲下限が正の整数15桁まで、少数2桁（半角数値で0.01から999999999999999.99までの入力）までであるか。
                If Not dc("下限").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("下限").ToString, "^[0-9]{1,15}(\.[0-9]{1,2})?$") Then
                        '---REV002 ADD START
                        '負の整数も許容する。
                        If Not Regex.IsMatch(dc("下限").ToString, "^-[0-9]{1,16}(\.[0-9]{1,2})?$") Then
                            '---REV002 ADD END
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                            '---REV002 ADD START
                        End If
                        '---REV002 ADD END
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("下限"), val) Then
                            '---REV002 MOD START
                            '負の整数を許容するので、絶対値にて判定
                            If Not (Math.Abs(val) >= 0.01D And Math.Abs(val) <= 999999999999999.99D) Then
                                'If Not (val >= 0.01D And val <= 999999999999999.99D) Then
                                '---REV002 MOD END
                                cnt = cnt + 1
                                details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If
                End If

                '6）データ範囲上限が正の整数15桁まで、少数2桁（半角数値で0.01から999999999999999.99までの入力）までであるか。
                If Not dc("上限").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("上限").ToString, "^[0-9]{1,15}(\.[0-9]{1,2})?$") Then
                        '---REV002 ADD START
                        '負の整数も許容する。
                        If Not Regex.IsMatch(dc("上限").ToString, "^-[0-9]{1,16}(\.[0-9]{1,2})?$") Then
                            '---REV002 ADD END
                            cnt = cnt + 1
                            details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                            '---REV002 ADD START
                        End If
                        '---REV002 ADD END
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("上限"), val) Then
                            '---REV002 MOD START
                            '負の整数を許容するので、絶対値にて判定
                            If Not (Math.Abs(val) >= 0.01D And Math.Abs(val) <= 999999999999999.99D) Then
                                'If Not (val >= 0.01D And val <= 999999999999999.99D) Then
                                '---REV002 MOD END
                                cnt = cnt + 1
                                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If
                End If

                '7）データ範囲上限 > データ範囲下限でがあるか。
                If Not (dc("上限").ToString.Equals(String.Empty) Or dc("下限").ToString.Equals(String.Empty)) Then
                    If Decimal.TryParse(dc("上限"), val2) And Decimal.TryParse(dc("下限"), val3) Then
                        If Not val2 > val3 Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    Else
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            Next

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 集計結果表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SyukeiKekkahyo

        ''' <summary>
        ''' 調査年コンボボックス設定
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="db"></param>
        ''' <param name="koutei"></param>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetChosaNenComboBox(chosaNen As ComboBox, db As DBAccess, koutei As String, chosakubun As String, kyoku As String, jimusho As String, kyoten As String)
            Dim dt As DataTable = Nothing
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOSyukeiKekkahyo.GetChosaNen(db, chosakubun, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOSyukeiKekkahyo.GetChosaNen(db, chosakubun, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOSyukeiKekkahyo.GetChosaNen(db, chosakubun, Nothing, Nothing, Nothing)
            End Select

            chosaNen.ValueMember = "調査年"
            chosaNen.DisplayMember = "調査年"
            chosaNen.DataSource = dt
        End Sub

        ''' <summary>
        ''' 集計結果表項目取得
        ''' </summary>
        ''' <param name="dtItem"></param>
        ''' <param name="dc"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetItem(dtItem As DataTable, dc As Dictionary(Of String, DataTable), Optional nousanFlg As Boolean = False) As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
            Dim ret As New Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)

            Dim lstItem As List(Of DataRow) = dtItem.AsEnumerable().ToList()

            For Each kv As KeyValuePair(Of String, DataTable) In dc
                Dim dt As DataTable = kv.Value
                Dim lstDr As List(Of DataRow) = dt.AsEnumerable().ToList()
                Dim colArr(dt.Columns.Count - 1) As DataColumn
                dt.Columns.CopyTo(colArr, 0)

                For Each row As DataRow In lstDr
                    For Each col As DataColumn In colArr
                        Dim query = (From dr In lstItem Where dr("項目番号").ToString = col.ColumnName Select dr).Take(1).ToArray
                        If query.Any() Then
                            Dim item As New DAOSyukeiKekkahyo.集計結果表項目
                            With item
                                .シート名 = query(0)("シート名").ToString
                                .行位置 = Integer.Parse(query(0)("行位置").ToString)
                                .列位置 = Integer.Parse(query(0)("列位置").ToString)
                                If dtItem.Columns.Contains("表示単位") Then
                                    .表示単位 = query(0)("表示単位").ToString
                                End If
                                'REV_015 START---------------
                                '農産物生産費の場合、制度受取金・積立金等項目名を設定する
                                If nousanFlg Then
                                    If query(0)("出力項目名").ToString <> "" Then
                                        .値 = query(0)("出力項目名").ToString
                                    ElseIf Left((row(col.ColumnName).ToString), 4) = ComConst.制度受取金積立金等項目.任意項目名称.調査票項番 Then
                                        .値 = ""
                                    Else
                                        .値 = row(col.ColumnName).ToString
                                    End If
                                Else
                                    .値 = row(col.ColumnName).ToString
                                End If
                                'REV_015 END---------------
                            End With
                            ret.Add(col.ColumnName, item)
                        End If
                    Next
                Next
            Next

            Return ret
        End Function

        ''' <summary>
        ''' 集計結果表シートデータ設定
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <param name="hiddenSheets">非表示化するシート名リスト(REV_017)</param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dc As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess, hiddenSheets As List(Of String))
            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    ' REV_017↓
                    '指定のシートを非表示化
                    If hiddenSheets.Count > 0 AndAlso hiddenSheets.Contains(sheet.シート名) Then
                        xlSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden
                        Continue For
                    End If
                    ' REV_017↑

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object

                        Dim page As Excel.PageSetup = xlSheet.PageSetup
                        Try
                            rng = xlSheet.Range(page.PrintArea)
                        Finally
                            ReleaseComObject(comObject, page)
                        End Try

                        arrData = DirectCast(rng.Formula, Object(,))

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 Select dr

                        For Each kv As KeyValuePair(Of String, DAOSyukeiKekkahyo.集計結果表項目) In query
                            Dim format As String = String.Empty

                            If Not String.IsNullOrEmpty(kv.Value.表示単位) Then
                                Dim unit As Decimal = Decimal.Parse(kv.Value.表示単位)
                                If ComConst.集計結果表作成論理.表示単位.リスト.ContainsKey(unit) Then
                                    format = ComConst.集計結果表作成論理.表示単位.リスト(unit)
                                End If
                            End If

                            Dim val As Decimal
                            If Not format.Equals(String.Empty) AndAlso Decimal.TryParse(kv.Value.値, val) Then
                                arrData(kv.Value.行位置, kv.Value.列位置) = val.ToString(format)
                            Else
                                arrData(kv.Value.行位置, kv.Value.列位置) = kv.Value.値
                            End If
                        Next

                        rng.Value = arrData
                        rng.Value = rng.Formula
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
        End Sub
    End Class

    ''' <summary>
    ''' 集計結果検討表クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SyukeiKekkaKentohyo

        ''' <summary>
        ''' 集計結果表項目クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 集計結果検討表項目
            Public シート名 As String                  'シート名
            Public 行位置 As Integer                   '行位置
            Public 列位置 As Integer                   '列位置
            Public 値 As String                        '値
            Public 型区分 As String                    '型区分
            Public 有効桁数 As Integer                 '有効桁数
            Public 小数点以下桁数 As Integer           '小数点以下桁数
            Public 表示単位 As String                  '表示単位
        End Class

        ''' <summary>
        ''' 集計結果検討表シートデータ設定
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dc As Dictionary(Of String, 集計結果検討表項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess)
            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object

                        Dim page As Excel.PageSetup = xlSheet.PageSetup
                        Try
                            rng = xlSheet.Range(page.PrintArea)
                        Finally
                            ReleaseComObject(comObject, page)
                        End Try

                        arrData = DirectCast(rng.Formula, Object(,))

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 Select dr

                        For Each kv As KeyValuePair(Of String, 集計結果検討表項目) In query
                            Dim format As String = String.Empty

                            If Not String.IsNullOrEmpty(kv.Value.表示単位) Then
                                Dim unit As Decimal = Decimal.Parse(kv.Value.表示単位)
                                If ComConst.集計結果検討表作成論理.表示単位.リスト.ContainsKey(unit) Then
                                    format = ComConst.集計結果検討表作成論理.表示単位.リスト(unit)
                                End If
                            End If

                            Dim val As Decimal
                            If Not format.Equals(String.Empty) AndAlso Decimal.TryParse(kv.Value.値, val) Then
                                arrData(kv.Value.行位置, kv.Value.列位置) = val.ToString(format)
                            Else
                                arrData(kv.Value.行位置, kv.Value.列位置) = kv.Value.値
                            End If
                        Next

                        rng.Value = arrData
                        rng.Value = rng.Formula
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
        End Sub
    End Class

    ''' <summary>
    ''' 集計結果表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SyukeiKekkahyoSakuseiRonri

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.集計結果表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.集計結果表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果表作成論理.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.集計結果表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果表作成論理.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.集計結果表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果表作成論理.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.集計結果表作成論理.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.集計結果表作成論理.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.集計結果表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.集計結果表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果表作成論理.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.集計結果表作成論理.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.集計結果表作成論理.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.集計結果表作成論理.出力用ファイル名称.Field
                            If kv.Value = "生産費平均値種類" Then
                                arrData(i, kv.Key) = ComConst.生産費平均値種類.リスト(dt.Rows(i - 1)(kv.Value).ToString)
                            Else
                                arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                            End If
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ' REV_013↓
        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        'Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), chosaKubun As String) As Boolean
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), chosaKubun As String, versionKbn As String) As Boolean
            ' REV_013↑
            Dim ret As Boolean = True
            Dim dtKobetsuItemMst As DataTable
            Dim dtSyukeiItemMst As DataTable

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　項番（{2}）の「生産費平均値種類」、「出力項目名」、「優先順位」、「計算式」は全て入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「出力項目名」、「優先順位」、「計算式」は全て入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」と「生産費平均値種類」の組み合わせがファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「生産費平均値種類」は「総数」か「総数以外」かで入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「優先順位」は半角数値0.1～99.9で入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」の条件式が不正です。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「表示単位」は半角数値で9、9.9、9.99、9.999のいずれかで入力してください。"
            }

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                '個別結果表項目マスタ取得
                '個別結果表項目マスタ取得(裏項番含める)
                ' REV_013
                If chosaKubun.Equals(ComConst.営農経営体区分.農業経営体) Then
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_個人, True)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_個人, versionKbn, True)
                    'dtKobetsuItemMst.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_法人, True))
                    dtKobetsuItemMst.Merge(DAOOther.GetKobetsuKekkahyoItemMaster(db, ComConst.調査区分.営農類型別経営統計_法人, versionKbn, True))
                Else
                    'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, chosaKubun, True)
                    dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, chosaKubun, versionKbn, True)
                End If
                '集計結果表項目マスタ取得(裏項番含める)
                'dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, chosaKubun)
                dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, chosaKubun, versionKbn)
                ' REV_013↑

                ''集計結果表作成クラス
                Dim syukei As CreateSyukei = New CreateSyukei(db,
                                                            chosaKubun,
                                                            Nothing,
                                                            dtKobetsuItemMst,
                                                            dtSyukeiItemMst,
                                                            Nothing,
                                                            True)

                Dim row As Integer = 0
                Dim cnt As Integer = 0
                '項番が存在するかのチェック用
                Dim isExistItemNo As Boolean

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    isExistItemNo = False

                    '1）①生産費平均値種類、出力項目名、優先順位、計算式、については、全て入力されているか。
                    If (CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 OrElse CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費) AndAlso
                       (dc("生産費平均値種類").ToString.Equals(String.Empty) OrElse
                        dc("項目名").ToString.Equals(String.Empty) OrElse
                        dc("優先順位").ToString.Equals(String.Empty) OrElse
                        dc("計算式").ToString.Equals(String.Empty)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '1）②出力項目名、優先順位、計算式、については、全て入力されているか。
                    If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 AndAlso
                       (dc("項目名").ToString.Equals(String.Empty) OrElse
                        dc("優先順位").ToString.Equals(String.Empty) OrElse
                        dc("計算式").ToString.Equals(String.Empty)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）項番が存在するか。
                    Dim query1 = From dr In dtSyukeiItemMst Where dr("項目番号").ToString = dc("項目番号").ToString Select dr
                    If Not query1.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        isExistItemNo = True
                    End If

                    '3）項番がファイル上で重複していないか。
                    Dim recCount As Integer
                    If (CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計) Then
                        recCount = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString).Count
                    Else
                        recCount = (From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString And dct("生産費平均値種類").ToString = dc("生産費平均値種類").ToString).Count
                    End If
                    If recCount > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '4）生産費平均値種類が「総数」か「総数以外」の入力であること。
                    If (CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 OrElse CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費) AndAlso
                       Not (dc("生産費平均値種類").Equals("総数") Or dc("生産費平均値種類").Equals("総数以外")) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '5）優先順位が半角数値で0.1から99.9までの入力であること。
                    If Not Regex.IsMatch(dc("優先順位").ToString, "^[0-9]{1,2}(\.[0-9]{1})?$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("優先順位"), val) Then
                            If Not (val >= 0.1 And val <= 99.9) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If

                    ''項番が存在する場合
                    If isExistItemNo Then
                        '6）計算式の入力が正しいかどうか。
                        'a)計算式として設定されている項番が存在するか
                        If syukei.CheckExistItemNo(dc("計算式").ToString, dtKobetsuItemMst, dtSyukeiItemMst, dc("生産費平均値種類").ToString) Then
                            'b)個別結果表作成を実行できる計算式となっているか。（SQLエラーとならない事をチェックする。）
                            If Not syukei.CheckExecutableSQL(dc("項目番号").ToString, dc("計算式").ToString, dtKobetsuItemMst, dtSyukeiItemMst, dc("生産費平均値種類").ToString) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        Else
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '7）表示単位が半角数値で9、9.9、9.99、9.999のいずれかの入力であること。
                    If Not dc("表示単位").ToString.Equals(String.Empty) Then
                        If Not Regex.IsMatch(dc("表示単位").ToString, "^[0-9]{1}(\.[0-9]{1,3})?$") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        Else
                            Dim val As Decimal
                            If Decimal.TryParse(dc("表示単位"), val) Then
                                If Not ComConst.集計結果表作成論理.表示単位.リスト.ContainsKey(val) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If
                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 集計結果検討表作成論理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SyukeiKekkaKentohyoSakuseiRonri

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.集計結果検討表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.集計結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.集計結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.集計結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.集計結果検討表作成論理.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.集計結果検討表作成論理.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.集計結果検討表作成論理.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.集計結果検討表作成論理.出力用ファイル名称.Col.First & ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.集計結果検討表作成論理.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.集計結果検討表作成論理.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.集計結果検討表作成論理.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ' REV_013 REV_014↓
        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <param name="chosaKubun"></param>
        ''' <param name="versionKbn"></param>
        ''' <param name="logicType"></param>
        ''' <returns></returns>
        'Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), ByVal chosaKubun As String) As Boolean
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String), ByVal chosaKubun As String, versionKbn As String, logicType As ComConst.集計結果検討表作成論理.論理種別) As Boolean
            ' REV_013 REV_014↑
            Dim ret As Boolean = True
            Dim dtSyukeiItemMst As DataTable
            Dim dtSyukeiKentoItemMst As DataTable

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                                 , "{0}件目：{1}行目　項番（{2}）の「出力項目名」、「優先順位」、「計算式」は全て入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「項番」がファイル上で重複しております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「優先順位」は半角数値0.1～99.9で入力してください。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」に存在しない項番が入力されております。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「計算式」の条件式が不正です。" _
                                 , "{0}件目：{1}行目　項番（{2}）の「表示単位」は半角数値で9、9.9、9.99、9.999のいずれかで入力してください。"
            }

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '集計結果表項目マスタ取得(裏項番含める)
                ' REV_013↓
                'dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, chosaKubun)
                dtSyukeiItemMst = DAOOther.GetSyukeiItemMaster(db, chosaKubun, versionKbn)
                ' REV_013↑
                '集計結果検討表項目マスタ取得
                ' REV_013 REV_014↓
                If logicType = ComConst.集計結果検討表作成論理.論理種別.集計結果検討表 Then
                    'dtSyukeiKentoItemMst = DAOOther.GetSyukeiKekkaKentohyoItemMaster(db, chosaKubun)
                    dtSyukeiKentoItemMst = DAOOther.GetSyukeiKekkaKentohyoItemMaster(db, chosaKubun, versionKbn)
                Else
                    dtSyukeiKentoItemMst = DAOOther.GetSyukeiKekkaKentohyoHoukokuyoItemMaster(db, chosaKubun, versionKbn)
                End If
                ' REV_013 REV_014↑

                '集計結果検討表作成クラス
                Dim kentohyo As CreateSyukeiKentohyo = New CreateSyukeiKentohyo(db,
                                                                                chosaKubun,
                                                                                dtSyukeiItemMst,
                                                                                dtSyukeiKentoItemMst,
                                                                                IsChikusan(),
                                                                                logicType) ' REV_016

                Dim row As Integer = 0
                Dim cnt As Integer = 0
                '項番が存在するかのチェック用
                Dim isExistItemNo As Boolean

                For Each dc As Dictionary(Of String, String) In lstDc
                    row = row + 1
                    isExistItemNo = False

                    '1）出力項目名、優先順位、計算式、については、全て入力されているか。
                    If dc("項目名").ToString.Equals(String.Empty) _
                        OrElse dc("優先順位").ToString.Equals(String.Empty) _
                        OrElse dc("計算式").ToString.Equals(String.Empty) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '2）項番が存在するか。
                    Dim query1 = From dr In dtSyukeiKentoItemMst Where dr("項目番号").ToString = dc("項目番号").ToString Select dr
                    If Not query1.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        isExistItemNo = True
                    End If

                    '3）項番がファイル上で重複していないか。
                    Dim query2 = From dct In lstDc Where dct("項目番号").ToString = dc("項目番号").ToString
                    If query2.Count > 1 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '4）優先順位が半角数値で0.1から99.9までの入力であること。
                    If Not Regex.IsMatch(dc("優先順位").ToString, "^[0-9]{1,2}(\.[0-9]{1})?$") Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                        ret = False
                        If cnt = max Then Return ret
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("優先順位"), val) Then
                            If Not (val >= 0.1 And val <= 99.9) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If

                    '項番が存在する場合
                    If isExistItemNo Then
                        '5）計算式の入力が正しいかどうか。
                        'a)計算式として設定されている項番が存在するか
                        If kentohyo.CheckExistItemNo(dc("項目番号").ToString, dc("計算式").ToString) Then
                            'b)集計結果検討表作成を実行できる計算式となっているか。（SQLエラーとならない事をチェックする。）
                            If Not kentohyo.CheckExecutableSQL(dc("項目番号").ToString, dc("計算式").ToString) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        Else
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If

                    '6）表示単位が半角数値で9、9.9、9.99、9.999のいずれかの入力であること。
                    If Not dc("表示単位").ToString.Equals(String.Empty) Then
                        If Not Regex.IsMatch(dc("表示単位").ToString, "^[0-9]{1}(\.[0-9]{1,3})?$") Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        Else
                            Dim val As Decimal
                            If Decimal.TryParse(dc("表示単位"), val) Then
                                If Not ComConst.集計結果検討表作成論理.表示単位.リスト.ContainsKey(val) Then
                                    cnt = cnt + 1
                                    details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4), dc("項目番号").ToString))
                                    ret = False
                                    If cnt = max Then Return ret
                                End If
                            End If
                        End If
                    End If
                Next

            End Using

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 任意帳票クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NiniChohyo

        ''' <summary>
        ''' 編成パラメータクラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 編成パラメータ
            Public タイトル As String
            Public 帳票レイアウト As String
            Public 開始位置セル As String
            Public 生産費平均値種類 As String
            Public 田畑区分指定 As String
            Public ビール麦販売区分 As String
            Public てんさい栽培区分 As String
            Public 表頭_表側 As String
            Public 項目番号 As List(Of String)
            Public 指標部 As List(Of 指標部)
        End Class

        ''' <summary>
        ''' 指標部クラス
        ''' </summary>
        ''' <remarks></remarks>
        Public Class 指標部
            Public 地域 As String
            Public 規模階層 As String
        End Class
    End Class

    ''' <summary>
    ''' 送受信管理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SendReceiveManagement

        ''' <summary>
        ''' 調査年コンボボックス設定（受信）
        ''' </summary>
        ''' <param name="chosaNen"></param>
        ''' <param name="db"></param>
        ''' <param name="koutei"></param>
        ''' <param name="upLow"></param>
        ''' <param name="dataType"></param>
        ''' <param name="kyoku"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="kyoten"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetChosaNenJushinComboBox(chosaNen As ComboBox, db As DBAccess, koutei As String, upLow As String, dataType As String, kyoku As String, jimusho As String, kyoten As String)
            Dim dt As DataTable = Nothing
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    dt = DAOOther.GetChosaNenJushin(db, upLow, dataType, kyoku, jimusho, kyoten)
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    dt = DAOOther.GetChosaNenJushin(db, upLow, dataType, kyoku, Nothing, Nothing)
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    dt = DAOOther.GetChosaNenJushin(db, upLow, dataType, Nothing, Nothing, Nothing)
            End Select

            chosaNen.ValueMember = "調査年"
            chosaNen.DisplayMember = "調査年"
            chosaNen.DataSource = dt
        End Sub

        ''' <summary>
        ''' 上位工程取得
        ''' </summary>
        ''' <param name="koutei"></param>
        ''' <param name="type"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetUpperKoutei(koutei As String) As String
            Dim ret As String = String.Empty

            '工程区分コード取得
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Center
                    ret = CommonInfo.KouteiKubun.Code.Kyoku
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    ret = CommonInfo.KouteiKubun.Code.Honsyo
            End Select

            Return ret
        End Function

        ''' <summary>
        ''' 下位工程取得
        ''' </summary>
        ''' <param name="koutei"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLowerKoutei(koutei As String) As String
            Dim ret As String = String.Empty

            '工程区分コード取得
            Select Case koutei
                Case CommonInfo.KouteiKubun.Code.Honsyo
                    ret = CommonInfo.KouteiKubun.Code.Kyoku
                Case CommonInfo.KouteiKubun.Code.Kyoku
                    ret = CommonInfo.KouteiKubun.Code.Center
            End Select

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 毎月勤労統計クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MaitsukiKinrouToukei

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="jimusho"></param>
        ''' <param name="xlSheet"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, jimusho As String, xlSheets As Excel.Sheets, comObject As ComObjectProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.毎月勤労統計.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '都道府県欄
                    rng = xlSheet.Range(ComConst.毎月勤労統計.出力用ファイル名称.Todofuken)
                    rng.Value = MasterDao.GetJimusyoName(jimusho)
                    ReleaseComObject(comObject, rng)

                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.毎月勤労統計.出力用ファイル名称.Col.First & ComConst.毎月勤労統計.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.毎月勤労統計.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.毎月勤労統計.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.毎月勤労統計.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula
                Finally
                    ReleaseComObject(comObject, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' 労賃単価クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RouchinTanka

        ''' <summary>
        ''' 対象期間取得
        ''' </summary>
        ''' <param name="seisanhi"></param>
        ''' <param name="chosaNen"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetTargetTerm(seisanhi As String, chosaNen As String) As Dictionary(Of String, String)
            Dim ret As New Dictionary(Of String, String)

            Select Case seisanhi
                Case ComConst.生産費区分.米_畜産物等_当年１月_当年12月
                    ret(ComConst.労賃単価.TERM_LOWER) = chosaNen & "01"
                    ret(ComConst.労賃単価.TERM_UPPER) = chosaNen & "12"
                Case ComConst.生産費区分.麦類_なたね_前年９月_当年８月
                    ret(ComConst.労賃単価.TERM_LOWER) = Integer.Parse(chosaNen) - 1 & "09"
                    ret(ComConst.労賃単価.TERM_UPPER) = chosaNen & "08"
                Case ComConst.生産費区分.さとうきび_当年４月_翌年３月
                    ret(ComConst.労賃単価.TERM_LOWER) = chosaNen & "04"
                    ret(ComConst.労賃単価.TERM_UPPER) = Integer.Parse(chosaNen) + 1 & "03"
                Case Else
                    Throw New Exception("生産費区分選択エラー")
            End Select

            Return ret
        End Function

        ''' <summary>
        ''' 生産費区分を取得する
        ''' </summary>
        ''' <param name="pChosaKubun"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSeisanhiKbn(ByVal pChosaKubun As String) As String
            Dim ret As String = Nothing

            Select Case (pChosaKubun)
                Case ComConst.調査区分.米生産費統計_個別,
                     ComConst.調査区分.そば生産費統計_個別,
                     ComConst.調査区分.大豆生産費統計_個別,
                     ComConst.調査区分.原料用かんしょ生産費統計_個別,
                     ComConst.調査区分.原料用ばれいしょ生産費統計_個別,
                     ComConst.調査区分.てんさい生産費統計_個別,
                     ComConst.調査区分.米生産費統計_組織法人,
                     ComConst.調査区分.大豆生産費統計_組織法人,
                     ComConst.調査区分.牛乳生産費統計_個別,
                     ComConst.調査区分.子牛生産費統計_個別,
                     ComConst.調査区分.乳用雄育成牛生産費統計_個別,
                     ComConst.調査区分.交雑種育成牛生産費統計_個別,
                     ComConst.調査区分.去勢若齢肥育牛生産費統計_個別,
                     ComConst.調査区分.乳用雄肥育牛生産費統計_個別,
                     ComConst.調査区分.交雑種肥育牛生産費統計_個別,
                     ComConst.調査区分.肥育豚生産費統計_個別,
                     ComConst.調査区分.経営分析調査_そば生産費,
                     ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費,
                     ComConst.調査区分.経営分析調査_てんさい生産費,
                     ComConst.調査区分.経営分析調査_牛乳生産費,
                     ComConst.調査区分.経営分析調査_子牛生産費,
                     ComConst.調査区分.経営分析調査_乳用雄育成牛生産費,
                     ComConst.調査区分.経営分析調査_交雑種育成牛生産費,
                     ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費,
                     ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費,
                     ComConst.調査区分.経営分析調査_交雑種肥育牛生産費,
                     ComConst.調査区分.経営分析調査_肥育豚生産費
                    ret = ComConst.生産費区分.米_畜産物等_当年１月_当年12月

                Case ComConst.調査区分.小麦生産費統計_個別,
                     ComConst.調査区分.二条大麦生産費統計_個別,
                     ComConst.調査区分.六条大麦生産費統計_個別,
                     ComConst.調査区分.はだか麦生産費統計_個別,
                     ComConst.調査区分.なたね生産費統計_個別,
                     ComConst.調査区分.小麦生産費統計_組織法人,
                     ComConst.調査区分.経営分析調査_二条大麦生産費,
                     ComConst.調査区分.経営分析調査_六条大麦生産費,
                     ComConst.調査区分.経営分析調査_はだか麦生産費,
                     ComConst.調査区分.経営分析調査_なたね生産費
                    ret = ComConst.生産費区分.麦類_なたね_前年９月_当年８月

                Case ComConst.調査区分.さとうきび生産費統計_個別,
                     ComConst.調査区分.経営分析調査_さとうきび生産費
                    ret = ComConst.生産費区分.さとうきび_当年４月_翌年３月
            End Select

            Return ret
        End Function

    End Class

    '---REV_003 Add Start
    ''' <summary>
    ''' 牛トレサクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Tresa

        ''' <summary>
        ''' 異動年月日の最終日とする年月日を取得
        ''' </summary>
        ''' <param name="chosaNen">調査年</param>
        ''' <param name="chosahyo">調査票データ</param>
        ''' <returns>異動年月日の最終日</returns>
        ''' <remarks></remarks>
        Public Shared Function GetIdoDateEnd(chosaNen As String, chosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As Integer
            Dim idoDate As Integer = Integer.Parse(chosaNen)
            Select Case CommonInfo.Chosakubun
                Case ComConst.調査区分.牛乳生産費統計_個別,
                    ComConst.調査区分.子牛生産費統計_個別,
                    ComConst.調査区分.乳用雄育成牛生産費統計_個別,
                    ComConst.調査区分.交雑種育成牛生産費統計_個別,
                    ComConst.調査区分.去勢若齢肥育牛生産費統計_個別,
                    ComConst.調査区分.乳用雄肥育牛生産費統計_個別,
                    ComConst.調査区分.交雑種肥育牛生産費統計_個別
                    '前年の12/31
                    idoDate -= 1
                    idoDate = (idoDate * 10000) + 1231
                Case ComConst.調査区分.経営分析調査_牛乳生産費,
                    ComConst.調査区分.経営分析調査_子牛生産費,
                    ComConst.調査区分.経営分析調査_乳用雄育成牛生産費,
                    ComConst.調査区分.経営分析調査_交雑種育成牛生産費,
                    ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費,
                    ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費,
                    ComConst.調査区分.経営分析調査_交雑種肥育牛生産費
                    '決算月取得
                    Dim month As Integer
                    If Not Integer.TryParse(chosahyo(ComConst.牛トレサデータ.決算月_項目番号(CommonInfo.Chosakubun)).値, month) Then
                        Return -1
                    End If

                    '決算月の末日（調査始め月の1日の前日）
                    Dim tmpDate As New Date(idoDate - 1, month, 1)
                    tmpDate = tmpDate.AddMonths(1)
                    tmpDate = tmpDate.AddDays(-1)
                    idoDate = (tmpDate.Year * 10000) + (tmpDate.Month * 100) + tmpDate.Day
            End Select

            Return idoDate
        End Function

        ''' <summary>
        ''' 調査開始日を返す
        ''' </summary>
        ''' <param name="endDate"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetChosaBeginDate(endDate As Integer) As Date
            Dim beginDate As New Date(endDate \ 10000, (endDate \ 100) Mod 100, endDate Mod 100)

            '終了日に対応する、1年前の開始日を求める　例：2020/4/30⇒2019/5/1
            beginDate = beginDate.AddDays(1)
            beginDate = beginDate.AddYears(-1)

            Return beginDate
        End Function


        ''' <summary>
        ''' 牛トレサデータが調査票への編成対象かどうかを返す
        ''' </summary>
        ''' <param name="kindCode">牛トレサデータ 牛の識別CD</param>
        ''' <param name="sexCode">牛トレサデータ 性別CD</param>
        ''' <returns>編成有無</returns>
        ''' <remarks></remarks>
        Public Shared Function IsOrganizeData(chosaBunrui As ComConst.牛トレサデータ.調査区分分類, kindCode As Integer, sexCode As Integer, idouFlag As Integer, Optional motherKotaiNo As String = "", Optional nyuyouList As List(Of String) = Nothing) As Boolean

            Dim ret As Boolean = False
            Dim isSex1_3_4 As Boolean = False

            '性別をめす／それ以外に丸める
            Select Case sexCode
                Case ComConst.牛トレサデータ.性別CD.おす,
                    ComConst.牛トレサデータ.性別CD.去勢,
                    ComConst.牛トレサデータ.性別CD.フリーマーチン
                    isSex1_3_4 = True
                Case ComConst.牛トレサデータ.性別CD.めす
                    '処理なし
                Case Else
                    '編成対象外
                    Return False
            End Select

            '異動フラグごとの編成有無判断
            Select Case idouFlag
                Case ComConst.牛トレサデータ.異動フラグ.出生,
                    ComConst.牛トレサデータ.異動フラグ.転出取引,
                    ComConst.牛トレサデータ.異動フラグ.転入搬入,
                    ComConst.牛トレサデータ.異動フラグ.死亡と畜
                    '処理なし
                Case Else
                    '編成対象外
                    Return False
            End Select

            '調査区分ごとの編成有無判断
            Select Case chosaBunrui
                Case ComConst.牛トレサデータ.調査区分分類.牛乳

                    Select Case kindCode
                        Case ComConst.牛トレサデータ.牛の識別CD.ホルスタイン種,
                                ComConst.牛トレサデータ.牛の識別CD.ジャージー種,
                                ComConst.牛トレサデータ.牛の識別CD.乳用種
                            If isSex1_3_4 Then
                                If Not nyuyouList Is Nothing Then
                                    If nyuyouList.Contains(motherKotaiNo) Then
                                        'めすでない場合、母牛が乳用種の場合のみ編成する
                                        ret = True
                                    End If
                                End If
                            Else
                                ret = True
                            End If
                        Case ComConst.牛トレサデータ.牛の識別CD.交雑種,
                            ComConst.牛トレサデータ.牛の識別CD.黒毛和種,
                            ComConst.牛トレサデータ.牛の識別CD.褐毛和種,
                            ComConst.牛トレサデータ.牛の識別CD.日本短角種,
                            ComConst.牛トレサデータ.牛の識別CD.無角和種
                            If Not nyuyouList Is Nothing Then
                                If nyuyouList.Contains(motherKotaiNo) Then
                                    '母牛が乳用種の場合のみ編成する
                                    ret = True
                                End If
                            End If
                    End Select

                Case ComConst.牛トレサデータ.調査区分分類.子牛

                    Select Case kindCode
                        Case ComConst.牛トレサデータ.牛の識別CD.黒毛和種,
                                ComConst.牛トレサデータ.牛の識別CD.褐毛和種,
                                ComConst.牛トレサデータ.牛の識別CD.日本短角種,
                                ComConst.牛トレサデータ.牛の識別CD.無角和種,
                                ComConst.牛トレサデータ.牛の識別CD.黒毛和種Ｘ褐毛和種
                            ret = True

                    End Select

                Case ComConst.牛トレサデータ.調査区分分類.乳用
                    Select Case kindCode
                        Case ComConst.牛トレサデータ.牛の識別CD.ホルスタイン種,
                            ComConst.牛トレサデータ.牛の識別CD.ジャージー種,
                            ComConst.牛トレサデータ.牛の識別CD.乳用種
                            If isSex1_3_4 Then
                                ret = True
                            End If
                    End Select

                Case ComConst.牛トレサデータ.調査区分分類.交雑

                    If kindCode = ComConst.牛トレサデータ.牛の識別CD.交雑種 Then
                        ret = True
                    End If

                Case ComConst.牛トレサデータ.調査区分分類.去勢

                    Select Case kindCode
                        Case ComConst.牛トレサデータ.牛の識別CD.黒毛和種,
                                ComConst.牛トレサデータ.牛の識別CD.褐毛和種,
                                ComConst.牛トレサデータ.牛の識別CD.日本短角種,
                                ComConst.牛トレサデータ.牛の識別CD.無角和種,
                                ComConst.牛トレサデータ.牛の識別CD.黒毛和種Ｘ褐毛和種
                            If isSex1_3_4 Then
                                ret = True

                            End If
                    End Select

            End Select

            Return ret
        End Function

        ''' <summary>
        ''' 調査票の種類コードを取得
        ''' </summary>
        ''' <param name="chosaBunrui"></param>
        ''' <param name="kindCode"></param>
        ''' <param name="sexCode"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetShuruiCode(chosaBunrui As ComConst.牛トレサデータ.調査区分分類, kindCode As Integer, sexCode As Integer, Optional isMakeAllIdo As Boolean = False) As String 'REV 023 CHG 
            Dim choShuruiCode As String

            If sexCode = ComConst.牛トレサデータ.性別CD.めす Then
                choShuruiCode = "2"

                If kindCode = ComConst.牛トレサデータ.牛の識別CD.交雑種 Then
                    '交雑種は1
                    choShuruiCode = "1"
                ElseIf chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
                    '牛乳生産費の場合、品種により異なる
                    Select Case kindCode
                        Case ComConst.牛トレサデータ.牛の識別CD.ホルスタイン種,
                            ComConst.牛トレサデータ.牛の識別CD.ジャージー種,
                            ComConst.牛トレサデータ.牛の識別CD.乳用種
                            '2のまま

                            'REV 023 ADD START --------------
                        Case ComConst.牛トレサデータ.牛の識別CD.黒毛和種,
                            ComConst.牛トレサデータ.牛の識別CD.褐毛和種,
                            ComConst.牛トレサデータ.牛の識別CD.日本短角種,
                            ComConst.牛トレサデータ.牛の識別CD.無角和種

                            choShuruiCode = "1"

                        Case Else
                            If isMakeAllIdo Then
                                '全所有牛情報出力の場合、編成対象外は2のまま
                            Else
                                choShuruiCode = "1"
                            End If
                            'REV 023 ADD END --------------
                    End Select
                End If
            Else
                choShuruiCode = "1"
            End If

            Return choShuruiCode
        End Function

        ''' <summary>
        ''' 調査票の品種コードを取得
        ''' </summary>
        ''' <param name="chosaBunrui"></param>
        ''' <param name="kindCode"></param>
        ''' <param name="sexCode"></param>
        ''' <param name="isMakeAllIdo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetHinshuCode(chosaBunrui As ComConst.牛トレサデータ.調査区分分類, kindCode As Integer, sexCode As Integer, Optional isMakeAllIdo As Boolean = False) As String 'REV 023 CHG 
            Dim tmpSexCode = sexCode
            Dim choHinsyuCode As String

            If sexCode <> ComConst.牛トレサデータ.性別CD.めす Then
                'Dictionaryの検索用におすに丸める
                tmpSexCode = ComConst.牛トレサデータ.性別CD.おす
            End If

            'REV 023 ADD START --------------
            '牛トレサデータプリント（異動データ）出力「全所有牛情報出力」の場合で品種コード変換テーブルに定義されていない品種について、品種コード99（その他）を設定する。
            If isMakeAllIdo And kindCode > ComConst.牛トレサデータ.牛の識別CD.黒毛和種Ｘ褐毛和種 And Not kindCode = ComConst.牛トレサデータ.牛の識別CD.乳用種 Then
                choHinsyuCode = "99"
            Else
                If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.牛乳 Then
                    choHinsyuCode = ComConst.牛トレサデータ.品種コード変換テーブル_牛乳(kindCode)(tmpSexCode)
                Else
                    choHinsyuCode = ComConst.牛トレサデータ.品種コード変換テーブル(kindCode)(tmpSexCode)
                End If
            End If
            'REV 023 ADD END --------------

            Return choHinsyuCode
        End Function

        ''' <summary>
        ''' 調査票の異動コードを取得
        ''' </summary>
        ''' <param name="chosaBunrui"></param>
        ''' <param name="idouFlag"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetIdouCode(chosaBunrui As ComConst.牛トレサデータ.調査区分分類, idouFlag As Integer) As String
            Dim choIdouCode As String

            If chosaBunrui = ComConst.牛トレサデータ.調査区分分類.子牛 Then
                choIdouCode = ComConst.牛トレサデータ.異動コード変換テーブル_子牛(idouFlag)
            Else
                choIdouCode = ComConst.牛トレサデータ.異動コード変換テーブル(idouFlag)
            End If

            Return choIdouCode
        End Function

        ''' <summary>
        ''' 成畜異動対象かを返す
        ''' </summary>
        ''' <param name="choShuruiCode"></param>
        ''' <param name="choHinsyuCode"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsAdultTarget(choShuruiCode As String, choHinsyuCode As String) As Boolean
            If choShuruiCode <> "2" Then
                Return False
            End If
            Select Case choHinsyuCode
                Case "1", "2", "3"
                    '処理なし(成畜異動対象対象)
                Case Else
                    Return False
            End Select

            Return True
        End Function

        ''' <summary>
        ''' 成畜年月日を取得する
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="cowID"></param>
        ''' <param name="birthday"></param>
        ''' <param name="baseDate"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAdultDate(db As DBAccess, cowID As String, birthday As String, baseDate As String) As String
            Dim childBirthday As String = GetFirstChildBirthday(db, cowID)

            Dim dBirth As New Date(
                Integer.Parse(birthday.Substring(0, 4)),
                Integer.Parse(birthday.Substring(4, 2)),
                Integer.Parse(birthday.Substring(6, 2)))
            Dim dBase As New Date(
                Integer.Parse(baseDate.Substring(0, 4)),
                Integer.Parse(baseDate.Substring(4, 2)),
                Integer.Parse(baseDate.Substring(6, 2)))
            Dim dChildBirth As Date

            If childBirthday IsNot Nothing Then
                dChildBirth = New Date(
                Integer.Parse(childBirthday.Substring(0, 4)),
                Integer.Parse(childBirthday.Substring(4, 2)),
                Integer.Parse(childBirthday.Substring(6, 2)))

                If dChildBirth <= dBirth.AddMonths(30) Then
                    '30か月齢以内に子を生産している場合、子の生産年月日
                    Return childBirthday
                End If
            End If

            If dBirth.AddMonths(30) < dBase Then
                '30か月齢を超えている場合、24か月齢に達した日
                Return dBirth.AddMonths(24).ToString("yyyyMMdd")
            End If

            Return ""
        End Function

        'REV 024 ADD START --------------
        ''' <summary>
        ''' 成畜年月日を取得する
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="cowID"></param>
        ''' <param name="birthday"></param>
        ''' <param name="baseDate"></param>
        ''' <param name="chousaDate"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAdultCattleDate(db As DBAccess, cowID As String, birthday As String, baseDate As String, chousaDate As String) As String
            Dim childBirthday As String = GetFirstChildBirthday(db, cowID)

            Dim dBirth As New Date(
                Integer.Parse(birthday.Substring(0, 4)),
                Integer.Parse(birthday.Substring(4, 2)),
                Integer.Parse(birthday.Substring(6, 2)))
            Dim dBase As New Date(
                Integer.Parse(baseDate.Substring(0, 4)),
                Integer.Parse(baseDate.Substring(4, 2)),
                Integer.Parse(baseDate.Substring(6, 2)))
            Dim dChildBirth As Date

            '調査時点の年月日
            Dim dChosa As Date = New Date(Integer.Parse(chousaDate), 1, 1)

            If childBirthday IsNot Nothing Then
                dChildBirth = New Date(
                Integer.Parse(childBirthday.Substring(0, 4)),
                Integer.Parse(childBirthday.Substring(4, 2)),
                Integer.Parse(childBirthday.Substring(6, 2)))

                If dBirth >= dChosa.AddMonths(-30) Then
                    '調査時点30か月齢以内に子牛を生産している場合、子牛の生産年月日
                    Return childBirthday
                End If
            End If

            Return ""
        End Function
        'REV 024 ADD END --------------


        ''' <summary>
        ''' 子牛の生年月日で、最も古いものの年月を取得
        ''' </summary>
        ''' <param name="db"></param>
        ''' <param name="cowID">個体識別番号</param>
        ''' <returns>子の生年月</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFirstChildBirthday(db As DBAccess, cowID As String) As String

            Dim birthday As String

            Dim dt As DataTable = DAOOther.GetTresaChild(db, cowID)

            If dt.Rows.Count = 0 Then
                birthday = Nothing
            Else
                Dim dbStr As String = dt.Rows(0)("生年月日").ToString
                If dbStr.Length = 8 Then
                    birthday = dbStr
                Else
                    birthday = Nothing
                End If
            End If

            Return birthday
        End Function


        ''' <summary>
        ''' 牛資産異動情報データの取得
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetDataIdoTresa(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess) As Dictionary(Of String, DAOChosahyo.調査票項目)
            Dim ret As New Dictionary(Of String, DAOChosahyo.調査票項目)

            Dim sheets = From dr In dt Group dr By dr!シート名 Into Group Select シート名
            Dim xlSheet As Excel.Worksheet = Nothing

            '20220201 ADD START 牛トレサの情報取得先取得時、旧区分と新区分で分けられるようにする
            'バージョン確認
            Dim ver_kubun As String        '1：2015年  2：2020年

            Dim xlVerCheckSheet As Excel.Worksheet = DirectCast(xlSheets.Item("表紙"), Excel.Worksheet)

            Dim xlVerCheckRange As Excel.Range = DirectCast(xlVerCheckSheet.Cells(13, 64), Excel.Range)
            '表紙の2015年センサス番号欄がセットされていたら2020年番、セット無しなら2015年番
            If Len(xlVerCheckRange.Value) = 0 Then
                ver_kubun = ComConst.バージョン区分.調査票項目2015
            Else
                ver_kubun = ComConst.バージョン区分.調査票項目2020
            End If
            '20220201 ADD END

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                    'シート保護確認
                    If xlSheet.ProtectContents Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object

                        rng = xlSheet.Range(ComConst.牛トレサデータ.シートデータ範囲(CommonInfo.Chosakubun)(sheet.ToString))

                        arrData = DirectCast(rng.Value, Object(,))

                        Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr

                        For Each dr As DataRow In query
                            '20220201 ADD START バージョン不一致レコード、行0列0のデータを読み込まないよう対応
                            'バージョン区分が一致しないレコードを飛ばす
                            If Not ver_kubun = dr("バージョン区分").ToString Then
                                Continue For
                            End If
                            '行0列0のデータは飛ばす
                            If String.IsNullOrEmpty(dr("行位置").ToString) OrElse dr("行位置").ToString = "0" OrElse
                                String.IsNullOrEmpty(dr("列位置").ToString) OrElse dr("列位置").ToString = "0" Then
                                Continue For
                            End If

                            '20220201 ADD END
                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                Dim item As New DAOChosahyo.調査票項目
                                With item
                                    .シート名 = dr("シート名").ToString
                                    .行位置 = Integer.Parse(dr("行位置").ToString)
                                    .列位置 = Integer.Parse(dr("列位置").ToString)
                                    .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                    .型区分 = dr("型区分").ToString
                                    .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                    .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                                End With
                                ret.Add(dr("項目番号").ToString, item)
                            Else
                                Dim increment As Integer = Integer.Parse(dr("可変増量").ToString)
                                For i As Integer = 1 To Integer.Parse(dr("可変最大数").ToString)
                                    Dim item As New DAOChosahyo.調査票項目
                                    With item
                                        .シート名 = dr("シート名").ToString
                                        .行位置 = Integer.Parse(dr("行位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                        .列位置 = Integer.Parse(dr("列位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        .型区分 = dr("型区分").ToString
                                    End With
                                    If Not String.IsNullOrEmpty(item.値) Then
                                        ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i.ToString, item)
                                    End If
                                Next
                            End If
                        Next
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
            Return ret
        End Function

        '---REV_004 Add Start
        ''' <summary>
        ''' 牛資産異動情報の登録対象のデータを取得
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetDataIdoTresaTarget(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess,
                                                          idoFrom As String, idoTo As String, kessan As String) As Dictionary(Of String, DAOChosahyo.調査票項目)
            Dim ret As New Dictionary(Of String, DAOChosahyo.調査票項目)

            Dim sheets = From dr In dt Group dr By dr!シート名 Into Group Select シート名
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                    'シート保護確認
                    If xlSheet.ProtectContents Then
                        xlSheet.Unprotect()
                    End If

                    '20220201 ADD START 牛トレサの情報取得先取得時、旧区分と新区分で分けられるようにする
                    'バージョン確認
                    Dim ver_kubun As String        '1：2015年  2：2020年

                    Dim xlVerCheckSheet As Excel.Worksheet = DirectCast(xlSheets.Item("表紙"), Excel.Worksheet)

                    Dim xlVerCheckRange As Excel.Range = DirectCast(xlVerCheckSheet.Cells(14, 64), Excel.Range)
                    '表紙の2015年センサス番号欄がセットされていたら2020年番、セット無しなら2015年番
                    If Len(xlVerCheckRange.Value) = 0 Then
                        ver_kubun = ComConst.バージョン区分.調査票項目2015
                    Else
                        ver_kubun = ComConst.バージョン区分.調査票項目2020
                    End If
                    '20220201 ADD END

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object

                        rng = xlSheet.Range(ComConst.牛トレサデータ.シートデータ範囲(CommonInfo.Chosakubun)(sheet.ToString))

                        arrData = DirectCast(rng.Value, Object(,))

                        Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr
                        Dim idotuki = From dr In dt Where dr("項目番号").ToString = ComConst.牛トレサデータ.異動月_項目番号(CommonInfo.Chosakubun)
                        Dim kessanNum As String = ComConst.牛トレサデータ.決算月_項目番号(CommonInfo.Chosakubun)

                        For Each dr As DataRow In query
                            '20220201 ADD START バージョン不一致レコード、行0列0のデータを読み込まないよう対応
                            'バージョン区分が一致しないレコードを飛ばす
                            If Not ver_kubun = dr("バージョン区分").ToString Then
                                Continue For
                            End If

                            '行0列0のデータは飛ばす
                            If String.IsNullOrEmpty(dr("行位置").ToString) OrElse dr("行位置").ToString = "0" OrElse
                                String.IsNullOrEmpty(dr("列位置").ToString) OrElse dr("列位置").ToString = "0" Then
                                Continue For
                            End If
                            '20220201 ADD END

                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                Dim item As New DAOChosahyo.調査票項目
                                With item
                                    .シート名 = dr("シート名").ToString
                                    .行位置 = Integer.Parse(dr("行位置").ToString)
                                    .列位置 = Integer.Parse(dr("列位置").ToString)
                                    .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                    .型区分 = dr("型区分").ToString
                                    .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                    .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                                End With
                                ret.Add(dr("項目番号").ToString, item)
                            Else
                                Dim increment As Integer = Integer.Parse(dr("可変増量").ToString)
                                Dim cntContinue As Integer = 0 'continue for した回数をカウント、登録時の行ずれに対応
                                For i As Integer = 1 To Integer.Parse(dr("可変最大数").ToString)
                                    Dim idoRow As Integer = Integer.Parse(idotuki(0)("行位置").ToString) + If(idotuki(0)("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                    Dim idoCol As Integer = Integer.Parse(idotuki(0)("列位置").ToString) + If(idotuki(0)("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                    Dim tmpIdoNengetsu As DateTime
                                    Dim idoTukiData As Integer = 0
                                    Dim idoTukiNengetsu As String

                                    '異動月の取得
                                    If IsNumeric(arrData(idoRow, idoCol)) Then
                                        idoTukiData = CInt(arrData(idoRow, idoCol))
                                    Else
                                        '異動月が取得できない場合、異動年月日から異動月を算出(牛乳のみ)
                                        If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
                                            Dim idoNengetu = From nen In dt Where nen("項目番号").ToString = ComConst.牛トレサデータ.異動年月_項目番号(CommonInfo.Chosakubun)
                                            Dim idoNenRow As Integer = Integer.Parse(idoNengetu(0)("行位置").ToString) + If(idoNengetu(0)("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                            Dim idoNenCol As Integer = Integer.Parse(idoNengetu(0)("列位置").ToString) + If(idoNengetu(0)("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                            '型と桁数をチェック
                                            If IsNumeric(arrData(idoNenRow, idoNenCol)) Then
                                                If arrData(idoNenRow, idoNenCol).ToString.Length = 1 Then
                                                    idoTukiData = 0
                                                ElseIf arrData(idoNenRow, idoNenCol).ToString.Length = 5 Then
                                                    tmpIdoNengetsu = DateTime.FromOADate(CDbl(arrData(idoNenRow, idoNenCol)))
                                                    idoTukiData = Month(tmpIdoNengetsu)
                                                End If
                                            End If
                                        End If
                                    End If

                                    '異動月、画面で入力した年月の範囲から、DB登録対象となるかを判定
                                    If idoTukiData <> 0 Then
                                        If CInt(kessan) < idoTukiData Then
                                            idoTukiNengetsu = CInt(ret(ComConst.調査票.項目番号.調査年).値.ToString) - 1 & idoTukiData.ToString.PadLeft(2, "0"c)
                                        Else
                                            idoTukiNengetsu = CInt(ret(ComConst.調査票.項目番号.調査年).値.ToString) & idoTukiData.ToString.PadLeft(2, "0"c)
                                        End If

                                        If Not (idoTukiNengetsu >= idoFrom And idoTukiNengetsu <= idoTo) Then
                                            '画面で入力した年月の範囲外であれば、DB登録の対象外とする。
                                            cntContinue = cntContinue + 1 'continue回数のカウント、明細行調整用、continueした分を登録時に詰める
                                            Continue For
                                        End If
                                    End If

                                    Dim item As New DAOChosahyo.調査票項目
                                    With item
                                        .シート名 = dr("シート名").ToString
                                        .行位置 = Integer.Parse(dr("行位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                        .列位置 = Integer.Parse(dr("列位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        .型区分 = dr("型区分").ToString
                                    End With
                                    If Not String.IsNullOrEmpty(item.値) Then
                                        ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & (i - cntContinue).ToString, item)
                                    End If
                                Next
                            End If
                        Next
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next
            Return ret
        End Function
        '---REV_004 Add END

        ''' <summary>
        ''' 電子調査票に牛トレサ関連データを設定する
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <param name="strChosaKubun"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dc As Dictionary(Of String, DAOChosahyo.調査票項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess, strChosaKubun As String)
            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            Dim sheetListIdo = ComConst.牛トレサデータ.異動ファイルシート(strChosaKubun)

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    '↓REV_021
                    '計算ボタン削除
                    If xlSheet.Name = sheetListIdo(1) Then
                        For Each shape As Excel.Shape In xlSheet.Shapes
                            If shape.AlternativeText = "計算" Then
                                shape.Delete()
                                Exit For
                            End If
                        Next
                    End If
                    '↑REV_021

                    '非可変項目
                    Dim cel As Excel.Range = Nothing
                    Try
                        cel = xlSheet.Cells

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 Is Nothing Select dr
                        For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query

                            Dim rng As Excel.Range = Nothing
                            Try
                                '2022/1/28 ADD START セル位置「0,0」を参照しないように修正 IF文追加
                                If (Not (kv.Value.行位置 = 0 Or kv.Value.列位置 = 0)) Then
                                    '2022/1/28 ADD END
                                    rng = DirectCast(cel.Item(kv.Value.行位置, kv.Value.列位置), Excel.Range)
                                    rng.Value = kv.Value.値
                                    '2022/1/28 ADD START セル位置「0,0」を参照しないように修正 IF文の対となるEndを追加
                                End If
                                '2022/1/28 ADD END
                            Finally
                                ReleaseComObject(comObject, rng)
                            End Try
                        Next

                    Finally
                        ReleaseComObject(comObject, cel)
                    End Try

                    '可変項目
                    Dim range As IEnumerable(Of String) = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 IsNot Nothing Select dr.Value.可変範囲 Distinct
                    For Each ar As String In range
                        Dim rng As Excel.Range = Nothing
                        Try
                            Dim arrData(,) As Object

                            rng = xlSheet.Range(ar)

                            arrData = DirectCast(rng.Formula, Object(,))

                            Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 = ar Select dr

                            For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query
                                arrData(kv.Value.行位置, kv.Value.列位置) = kv.Value.値
                            Next

                            rng.Value = arrData
                            rng.Value = rng.Formula
                        Finally
                            ReleaseComObject(comObject, rng)
                        End Try
                    Next

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next

            '不要なシートの削除
            Dim deleleSheetNames As New List(Of String)
            For Each xlSheet In xlSheets
                If Not ComConst.牛トレサデータ.異動ファイルシート(strChosaKubun).Contains(xlSheet.Name) Then
                    deleleSheetNames.Add(xlSheet.Name)
                End If
            Next
            If deleleSheetNames.Count <> 0 Then
                For Each sheetName As String In deleleSheetNames
                    If sheetName.Equals("リスト") Then
                        'リストシートは削除しない(非表示のまま)
                    Else
                        xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                        xlSheet.Delete()
                        ReleaseComObject(comObject, xlSheet)
                        xlSheet = Nothing
                    End If
                Next
            End If

        End Sub

        'REV-018 START-------------------
        ''' <summary>
        ''' 電子調査票に牛トレサ関連データを設定する
        ''' </summary>
        ''' <param name="dc"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <param name="strChosaKubun"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetDataSoukatsu(dc As Dictionary(Of String, DAOChosahyo.調査票項目), xlSheets As Excel.Sheets, comObject As ComObjectProcess, strChosaKubun As String)
            Dim sheets = From dr In dc Group dr By dr.Value.シート名 Into Group
            Dim xlSheet As Excel.Worksheet = Nothing

            Dim sheetListIdo = ComConst.牛トレサデータ.総括ファイルシート(strChosaKubun)

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet.シート名), Excel.Worksheet)

                    'シート保護確認
                    Dim protect As Boolean = xlSheet.ProtectContents
                    If protect Then
                        xlSheet.Unprotect()
                    End If

                    '↓REV_021
                    '計算ボタン削除
                    If xlSheet.Name = sheetListIdo(1) Then
                        For Each shape As Excel.Shape In xlSheet.Shapes
                            If shape.AlternativeText = "計算" Then
                                shape.Delete()
                                Exit For
                            End If
                        Next
                    End If
                    '↑REV_021

                    '非可変項目
                    Dim cel As Excel.Range = Nothing
                    Try
                        cel = xlSheet.Cells

                        Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 Is Nothing Select dr
                        For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query

                            Dim rng As Excel.Range = Nothing
                            Try
                                '2022/1/28 ADD START セル位置「0,0」を参照しないように修正 IF文追加
                                If (Not (kv.Value.行位置 = 0 Or kv.Value.列位置 = 0)) Then
                                    '2022/1/28 ADD END
                                    rng = DirectCast(cel.Item(kv.Value.行位置, kv.Value.列位置), Excel.Range)
                                    rng.Value = kv.Value.値
                                    '2022/1/28 ADD START セル位置「0,0」を参照しないように修正 IF文の対となるEndを追加
                                End If
                                '2022/1/28 ADD END
                            Finally
                                ReleaseComObject(comObject, rng)
                            End Try
                        Next

                    Finally
                        ReleaseComObject(comObject, cel)
                    End Try

                    '可変項目
                    Dim range As IEnumerable(Of String) = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 IsNot Nothing Select dr.Value.可変範囲 Distinct
                    For Each ar As String In range
                        Dim rng As Excel.Range = Nothing
                        Try
                            Dim arrData(,) As Object

                            rng = xlSheet.Range(ar)

                            arrData = DirectCast(rng.Formula, Object(,))

                            Dim query = From dr In dc Where dr.Value.シート名 = sheet.シート名 And dr.Value.可変範囲 = ar Select dr

                            For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In query
                                arrData(kv.Value.行位置, kv.Value.列位置) = kv.Value.値
                            Next

                            rng.Value = arrData
                            rng.Value = rng.Formula
                        Finally
                            ReleaseComObject(comObject, rng)
                        End Try
                    Next

                    If protect Then
                        xlSheet.Protect()
                    End If
                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next

            '不要なシートの削除
            Dim deleleSheetNames As New List(Of String)
            For Each xlSheet In xlSheets
                If Not ComConst.牛トレサデータ.総括ファイルシート(strChosaKubun).Contains(xlSheet.Name) Then
                    deleleSheetNames.Add(xlSheet.Name)
                End If
            Next
            If deleleSheetNames.Count <> 0 Then
                For Each sheetName As String In deleleSheetNames
                    If sheetName.Equals("リスト") Then
                        'リストシートは削除しない(非表示のまま)
                    Else
                        xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                        xlSheet.Delete()
                        ReleaseComObject(comObject, xlSheet)
                        xlSheet = Nothing
                    End If
                Next
            End If

        End Sub
        'REV-018 END-------------------

        ''' <summary>
        ''' 調査票データに含まれる農家団体コードをすべて取得
        ''' </summary>
        ''' <param name="chosahyo">調査票データ</param>
        ''' <returns>農家団体コードリスト</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFarmCodeList(chosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As List(Of String)
            Dim farmCodeList As New List(Of String)
            Dim farmCodeKey As String
            Dim i As Integer = 1
            While True
                farmCodeKey = ComConst.調査票.項目番号.牛農家団体コード(CommonInfo.Chosakubun) & ComConst.ITEM_NO_DELIMITER & i.ToString

                If chosahyo.ContainsKey(farmCodeKey) Then
                    farmCodeList.Add(chosahyo(farmCodeKey).値)
                Else
                    Exit While
                End If

                i += 1
            End While

            Return farmCodeList
        End Function

        ''' <summary>
        ''' 調査票の牛資産総括データの乳用種をすべて取得
        ''' </summary>
        ''' <param name="chosahyo">調査票データ</param>
        ''' <returns>個体識別番号リスト</returns>
        ''' <remarks></remarks>
        Public Shared Function GetNyuyouList(chosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)) As List(Of String)

            Dim kobanShurui As String = If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022001", "Q11021901")
            Dim kobanHinshu As String = If(CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別, "Q11022101", "Q11022001")
            Dim kobanKotaiNo As String = "Q11021801"

            Dim kotaiNoList As New List(Of String)

            If CommonInfo.Chosakubun <> ComConst.調査区分.牛乳生産費統計_個別 And
                CommonInfo.Chosakubun <> ComConst.調査区分.経営分析調査_牛乳生産費 Then
                Return kotaiNoList
            End If

            '対象項目を抽出
            Dim query1 = chosahyo.Where(Function(row) row.Key.StartsWith(kobanShurui) OrElse
                                                  row.Key.StartsWith(kobanHinshu) OrElse
                                                  row.Key.StartsWith(kobanKotaiNo)).ToList()

            '種類コードが2のレコードを抽出
            Dim queryType1 = query1.Where(Function(row) row.Key.StartsWith(kobanShurui) AndAlso
                                            row.Value.値 = "2").
                                Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))
            '品種コードが1～3のレコードを抽出
            Dim queryType2 = query1.Where(Function(row) row.Key.StartsWith(kobanHinshu) AndAlso
                                              (row.Value.値 = "1" OrElse row.Value.値 = "2" OrElse row.Value.値 = "3")).
                                Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))

            '個体識別番号を抽出
            Dim query2 = query1.Where(Function(row) row.Key.StartsWith(kobanKotaiNo)).
                                       Select(Function(row) ComUtil.Chosahyo.GetEdaNo(row.Key))
            query2 = query2.Intersect(queryType1.Intersect(queryType2))

            For Each edaNo In query2
                If chosahyo.ContainsKey(kobanKotaiNo & edaNo) Then
                    kotaiNoList.Add(chosahyo(kobanKotaiNo & edaNo).値)
                End If
            Next

            Return kotaiNoList
        End Function

    End Class

    Public Class OrganiezCheck
        ''' <summary>
        ''' 牛の異動状態
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum CowState
            STATE_IN
            STATE_OUT
            STATE_END
        End Enum

        ''' <summary>
        ''' 牛の異動状態による編成有無
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared OrganiezTable As New Dictionary(Of CowState, Boolean) From {
            {CowState.STATE_IN, True},
            {CowState.STATE_OUT, False},
            {CowState.STATE_END, False}
        }

        ''' <summary>
        ''' 牛トレサデータの異動フラグから、異動状態への変換テーブル
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared CowStateTable As New Dictionary(Of Integer, CowState) From {
            {ComConst.牛トレサデータ.異動フラグ.初期装着, CowState.STATE_IN},
            {ComConst.牛トレサデータ.異動フラグ.出生, CowState.STATE_IN},
            {ComConst.牛トレサデータ.異動フラグ.転出取引, CowState.STATE_OUT},
            {ComConst.牛トレサデータ.異動フラグ.転入搬入, CowState.STATE_IN},
            {ComConst.牛トレサデータ.異動フラグ.死亡と畜, CowState.STATE_END},
            {ComConst.牛トレサデータ.異動フラグ.と場, CowState.STATE_END},
            {ComConst.牛トレサデータ.異動フラグ.輸出, CowState.STATE_OUT},
            {ComConst.牛トレサデータ.異動フラグ.輸入, CowState.STATE_IN},
            {ComConst.牛トレサデータ.異動フラグ.とさつ, CowState.STATE_END}
       }

        ''' <summary>
        ''' 異動フラグ
        ''' </summary>
        ''' <param name="row"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetCowStateFromRow(row As DataRow) As CowState
            Dim state As CowState = CowState.STATE_END

            Dim flag As Integer = Integer.Parse(row("異動フラグ").ToString)

            If CowStateTable.ContainsKey(flag) Then
                state = CowStateTable(flag)
            End If
            Return state
        End Function

        ''' <summary>
        ''' 牛トレサ情報の異動履歴をトレースし、当該客体に牛が存在しているかを返す
        ''' </summary>
        ''' <param name="dtTresaRows">牛トレサ情報（同一個体識別番号のみ、異動年月日降順で整列済みの前提）</param>
        ''' <returns>牛の存在有無</returns>
        ''' <remarks></remarks>
        Public Shared Function IsCowOrganize(dtTresaRows As DataRow(), ByRef tennyuIndex As Integer) As Boolean

            Dim isOrgan As Boolean = False
            tennyuIndex = -1

            For i As Integer = 0 To dtTresaRows.Count - 1
                If i >= dtTresaRows.Count Then
                    '牛のIN情報がないため、編成対象外
                    Exit For
                End If

                If Integer.Parse(dtTresaRows(i)("異動フラグ").ToString) = ComConst.牛トレサデータ.異動フラグ.転入搬入 Then
                    '取得年月日を取得する際のインデックスを保持
                    tennyuIndex = If(tennyuIndex >= 0, tennyuIndex, i)
                End If

                If i >= dtTresaRows.Count - 1 Then
                    '最終レコードの場合、当該レコードにより編成有無を判断
                    isOrgan = ComUtil.OrganiezCheck.OrganiezTable(ComUtil.OrganiezCheck.GetCowStateFromRow(dtTresaRows(i)))
                    Exit For
                End If

                If dtTresaRows(i)("異動年月日").ToString = dtTresaRows(i + 1)("異動年月日").ToString Then
                    '異動年月日が同一のレコードが2件ある場合
                    Dim state1 As CowState = ComUtil.OrganiezCheck.GetCowStateFromRow(dtTresaRows(i))
                    Dim state2 As CowState = ComUtil.OrganiezCheck.GetCowStateFromRow(dtTresaRows(i + 1))

                    If state1 = CowState.STATE_END Or state2 = CowState.STATE_END Then
                        'どちらかがENDの場合、編成対象外
                        Exit For
                    End If

                    If Integer.Parse(dtTresaRows(i + 1)("異動フラグ").ToString) = ComConst.牛トレサデータ.異動フラグ.転入搬入 Then
                        '取得年月日を取得する際のインデックスを保持
                        tennyuIndex = If(tennyuIndex >= 0, tennyuIndex, i + 1)
                    End If

                    If (state1 = CowState.STATE_IN And state2 = CowState.STATE_OUT) Or
                        (state2 = CowState.STATE_IN And state1 = CowState.STATE_OUT) Then
                        '2つのレコードがIN/OUTの組合せの場合、存在有無の変更がないとして、2つ(Forの加算含む)過去のレコードを見る
                        i += 1
                    Else
                        '2つのレコードがIN、OUTどちらかのみの場合、当該レコードにより編成有無を判断
                        isOrgan = OrganiezTable(state1)
                        Exit For
                    End If
                Else
                    '異動年月日が同一でない場合
                    isOrgan = OrganiezTable(GetCowStateFromRow(dtTresaRows(i)))
                    Exit For
                End If
            Next

            If tennyuIndex < 0 Then
                tennyuIndex = 0
            End If


            Return isOrgan
        End Function
    End Class
    '---REV_003 Add End


    ''' <summary>
    ''' ファイルパス取得時(Excel UserForm表示時)のコールバック
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <remarks>REV007 Add</remarks>
    Public Delegate Sub FileDialogWithNewThreadCallback(filePath As String)

    ''' <summary>
    ''' ファイルパス取得用クラス(Excel UserForm表示時)
    ''' </summary>
    ''' <remarks>REV007 Add</remarks>
    Public Class fileDialogWithNewThread
        Private _initialDirectory As String
        Private _fileName As String

        Private _callback As FileDialogWithNewThreadCallback

        Public Sub New(initialDirectory As String, fileName As String, callback As FileDialogWithNewThreadCallback)
            _fileName = fileName
            _initialDirectory = initialDirectory
            _callback = callback
        End Sub

        Public Sub GetFilePath()

            Dim ret As String = String.Empty
            Dim fdg As New SaveFileDialog()

            fdg.InitialDirectory = _initialDirectory
            'fdg.Filter = Filter()
            fdg.FilterIndex = 1
            fdg.RestoreDirectory = True

            If Not _fileName Is Nothing Then
                fdg.FileName = _fileName
            End If

            Dim tp As Type = fdg.GetType()
            Dim pInf As Reflection.PropertyInfo = tp.GetProperty("OverwritePrompt")
            pInf.SetValue(fdg, False)

            If Not System.IO.Directory.Exists(_initialDirectory) Then
                System.IO.Directory.CreateDirectory(_initialDirectory)
            End If

            If fdg.ShowDialog() = DialogResult.OK Then
                ret = fdg.FileName
            End If

            _callback(ret)

        End Sub
    End Class

    ''' <summary>
    ''' パラメーターの値を合計する
    ''' </summary>
    ''' <param name="values"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Sum(ParamArray values As Decimal()) As Decimal

        If values Is Nothing OrElse values.Length = 0 Then
            Return 0
        End If

        Dim result = CDec(0)

        For Each value In values
            result += value
        Next

        Return result

    End Function

    ''' <summary>
    ''' 指定した精度の数値に四捨五入する
    ''' </summary>
    ''' <param name="value">値</param>
    ''' <param name="iDigits">戻り値の有効桁数の精度</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Round(ByVal value As Decimal, Optional ByVal iDigits As Integer = 0) As Decimal
        Dim pow As Decimal
        Dim ret As Decimal

        Try
            pow = CDec(System.Math.Pow(10, iDigits))
            ret = Math.Round(value * pow, MidpointRounding.AwayFromZero) / pow
        Catch ex As Exception
            Throw ex
        End Try

        Return ret

    End Function

    ''' <summary>
    ''' 引数がDBNullの場合文字列の空文字列を返す
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetBlankOrString(ByVal obj As Object) As String

        If obj Is DBNull.Value Then
            Return ""
        Else
            Return obj.ToString
        End If

    End Function

    ''' <summary>
    ''' 引数がDBNullの場合Nothingを返す
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns>
    ''' TRUE：Nothing
    ''' FALSE：引数をIntegerに変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNullOrInt(ByVal obj As Object) As Integer?

        If IsDBNull(obj) Then
            Return Nothing
        End If

        Return CInt(obj)

    End Function

    ''' <summary>
    ''' 引数がDBNullの場合Nothingを返す
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns>
    ''' TRUE：Nothing
    ''' FALSE：引数をDecimalに変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNullOrDec(ByVal obj As Object) As Decimal?

        If IsDBNull(obj) Then
            Return Nothing
        End If

        Return CDec(obj)

    End Function

    ''' <summary>
    ''' 引数がDBNullの場合Nothingを返す
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns>
    ''' TRUE：Nothing
    ''' FALSE：引数をDateに変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function GetNullOrDateConvertString(ByVal obj As Object) As Date?

        If IsDBNull(obj) Then
            Return Nothing
        End If

        Return Date.Parse(obj.ToString())

    End Function

    ''' <summary>
    ''' 引数がDBNullの場合0を返す
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns>
    ''' TRUE：0
    ''' FALSE：引数をDecimalに変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function GetZeroOrDec(ByVal obj As Object) As Decimal

        If IsDBNull(obj) Then
            Return 0
        End If

        Return CDec(obj)

    End Function

    ''' <summary>
    ''' 引数が数値に変換可能な文字列の場合はDecimalに変換して返し、変換不可の場合は0を返す
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns>
    ''' TRUE：0
    ''' FALSE：引数をDecimalに変換した値</returns>
    ''' <remarks></remarks>
    Public Shared Function GetZeroOrDec(ByVal value As String) As Decimal

        If Not TryParseToDecimal(value) Then
            Return 0
        End If

        Return CDec(value)

    End Function

    ''' <summary>
    ''' 文字列を数値型に変換できるかチェックする
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>
    ''' True ：変換できる
    ''' False：変換できない
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function TryParseToInteger(ByVal value As String) As Boolean

        Dim buf As Integer = 0

        Return Integer.TryParse(value, buf)

    End Function

    ''' <summary>
    ''' 文字列を数値型に変換できるかチェックする
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>
    ''' True ：変換できる
    ''' False：変換できない
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function TryParseToDecimal(ByVal value As String) As Boolean

        Dim buf As Decimal = 0

        Return Decimal.TryParse(value, buf)

    End Function

    ' REV_013↓
    ''' <summary>
    ''' 【集計結果表作成】固定文字列の検索条件(任意階層、集計条件１～４)を該当する調査区分の項番へ置換する
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function ReplaceConstJouken(ByVal value As String, ByVal chosakubun As String) As String
    Public Shared Function ReplaceConstJouken(ByVal value As String, ByVal chosakubun As String, versionKbn As String) As String
        ' REV_013↑
        Dim ret As String = value

        For Each kv As KeyValuePair(Of String, Dictionary(Of String, String)) In ComConst.個別結果表.任意集計条件文字列(versionKbn)
            If ComConst.個別結果表.任意集計条件文字列(versionKbn)(kv.Key).ContainsKey(chosakubun) Then
                ret = ret.Replace(kv.Key, "ISNULL(" & kv.Value.Item(chosakubun) & ", 0)")
            End If
        Next

        Return ret
    End Function

    ' REV_013↓
    ''' <summary>
    ''' 【集計結果表作成】固定文字列の検索条件(任意階層、集計条件１～４)を該当する調査区分の項番へ置換する
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="versionKbn"></param>
    ''' <returns></returns>
    'Public Shared Function CheckConstJouken(ByVal value As String, ByVal chosakubun As String) As Boolean
    Public Shared Function CheckConstJouken(ByVal value As String, ByVal chosakubun As String, versionKbn As String) As Boolean
        ' REV_013↑
        Dim ret As Boolean = True

        ' REV_013↓
        'If Not ComConst.個別結果表.任意集計条件文字列.ContainsKey(value) Then
        If Not ComConst.個別結果表.任意集計条件文字列(versionKbn).ContainsKey(value) Then
            '該当の文字列が置換文字列に定義されているか
            ret = False

            'ElseIf Not ComConst.個別結果表.任意集計条件文字列(value).ContainsKey(chosakubun) Then
        ElseIf Not ComConst.個別結果表.任意集計条件文字列(versionKbn)(value).ContainsKey(chosakubun) Then
            '該当の置換文字列が指定した調査区分に存在するか
            ret = False
        End If
        ' REV_013↑

        Return ret
    End Function

    ''' <summary>
    ''' 調査区分が畜産かどうかを返す
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsChikusan() As Boolean
        Dim ret As Boolean = False

        If CommonInfo.Kubun2.Equals(ComConst.区分２.畜産物生産費) Then
            ret = True
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 調査区分が営農かどうかを返す
    ''' </summary>
    ''' <param name="chosakubun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsEinou() As Boolean
        Dim ret As Boolean = False

        If CommonInfo.Kubun2.Equals(ComConst.区分２.営農類型別経営統計) Then
            ret = True
        End If

        Return ret
    End Function

    ' REV_016↓
    ''' <summary>
    ''' 調査区分が農産かどうかを返す
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function IsNousan() As Boolean
        Dim ret As Boolean = False

        If CommonInfo.Kubun2.Equals(ComConst.区分２.農産物生産費) Then
            ret = True
        End If

        Return ret
    End Function
    ' REV_016↑

    ''' <summary>
    ''' 営農類型名に「経営」をつけて返す
    ''' </summary>
    ''' <param name="kobetsu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getEinouName(ByVal kobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)) As String
        Dim ret As String = String.Empty

        Dim einou As String = kobetsu(ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)).値

        ret = ComConst.営農類型区分.リスト(einou) & "経営"

        Return ret
    End Function


    'REV-005 START------------------------------------------
    ''' <summary>
    ''' 調査年に対してバージョン区分を返す
    ''' </summary>
    ''' <param name="kobetsu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getVersionKubun(chosanen As String, chosakubun As String) As String

        Dim ver_kubun As String = String.Empty

        If ((CInt(chosanen)) >= 2022 And Not (chosakubun = ComConst.調査区分.小麦生産費統計_個別 Or
                                          chosakubun = ComConst.調査区分.小麦生産費統計_組織法人 Or
                                          chosakubun = ComConst.調査区分.二条大麦生産費統計_個別 Or
                                          chosakubun = ComConst.調査区分.経営分析調査_二条大麦生産費 Or
                                          chosakubun = ComConst.調査区分.六条大麦生産費統計_個別 Or
                                          chosakubun = ComConst.調査区分.経営分析調査_六条大麦生産費 Or
                                          chosakubun = ComConst.調査区分.はだか麦生産費統計_個別 Or
                                          chosakubun = ComConst.調査区分.経営分析調査_はだか麦生産費 Or
                                          chosakubun = ComConst.調査区分.なたね生産費統計_個別 Or
                                          chosakubun = ComConst.調査区分.経営分析調査_なたね生産費)) Then


            ver_kubun = ComConst.バージョン区分.調査票項目2020
        ElseIf ((CInt(chosanen)) >= 2023) Then
            ver_kubun = ComConst.バージョン区分.調査票項目2020
        Else
            ver_kubun = ComConst.バージョン区分.調査票項目2015
        End If

        Return ver_kubun
    End Function
    'REV-005 END------------------------------------------

    ' REV_011↓
    ''' <summary>
    ''' 調査年に対してバージョン区分を返す
    ''' </summary>
    ''' <param name="chosanen"></param>
    ''' <returns></returns>
    Public Shared Function getVersionKubunTaikei(chosanen As String, chosakubun As String) As String
        Dim ver_kubun As String = String.Empty

        If (CInt(chosanen) >= 2022 And ComConst.令和４年体系.対象調査区分2022.IndexOf(chosakubun) > -1) _
            Or (CInt(chosanen) >= 2023 And ComConst.令和４年体系.対象調査区分2023.IndexOf(chosakubun) > -1) Then
            ver_kubun = ComConst.バージョン区分.結果表等項目2022
        Else
            ver_kubun = ComConst.バージョン区分.結果表等項目2021
        End If

        Return ver_kubun
    End Function
    ' REV_011↑

    ''' <summary>
    ''' ロックファイルを削除し、ストリームを解放する
    ''' </summary>
    ''' <param name="lockFileStream"></param>
    ''' <remarks>REV-008 ADD</remarks>
    Public Shared Sub DeleteLockFile(ByRef lockFileStream As System.IO.FileStream)
        Dim lockFilePath As String = lockFileStream.Name

        lockFileStream.Dispose()
        lockFileStream = Nothing

        System.IO.File.Delete(lockFilePath)

    End Sub

    ''' <summary>
    ''' ロックファイルのパスを取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>REV-008 ADD</remarks>
    Public Shared Function GetLockFilePath(chosaNen As String, censusNo As String) As String
        Dim lockFileName = "調査票_{0}_電子調査票_{1}_{2}.lock"
        Return IniFileInfo.LockPath & "\" & String.Format(lockFileName, ComUtil.GetChosakubunName(CommonInfo.Chosakubun), chosaNen, censusNo)
    End Function

    ''' <summary>
    ''' テンプレートファイル(コピー)のパスを取得
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="censusNo"></param>
    ''' <param name="isKoumoku"></param>
    ''' <returns></returns>
    Public Shared Function GetTemplateCopyFilePath(chosaNen As String, censusNo As String, Optional isKoumoku As Boolean = False) As String
        Dim templateFileName = ComConst.調査票.入力用ファイル名称(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(chosaNen, CommonInfo.Chosakubun)))
        Dim templateCopyFileName As String
        If isKoumoku Then
            templateCopyFileName = "CC_{0}_{1}_{2}"
        Else
            templateCopyFileName = "C_{0}_{1}_{2}"
        End If

        Return System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), String.Format(templateCopyFileName, chosaNen, censusNo, templateFileName))
    End Function

    ''' <summary>
    ''' 調査票審査論理範囲クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ChosahyoShinsaRonriRange

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.調査票審査論理範囲.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.調査票審査論理範囲.出力用ファイル名称.Col.First & ComConst.調査票審査論理範囲.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.調査票審査論理範囲.出力用ファイル名称.Col.First & ComConst.調査票審査論理範囲.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.調査票審査論理範囲.出力用ファイル名称.Col.First & ComConst.調査票審査論理範囲.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.調査票審査論理範囲.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.調査票審査論理範囲.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            dc("連番") = i.ToString
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.調査票審査論理範囲.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.調査票審査論理範囲.出力用ファイル名称.Col.First & ComConst.調査票審査論理範囲.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.調査票審査論理範囲.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.調査票審査論理範囲.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.調査票審査論理範囲.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula
                Finally
                    ReleaseComObject(comObject, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try
        End Sub

        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String)) As Boolean
            Dim ret As Boolean = True

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim val1 As Long
            Dim val2 As Decimal
            Dim val3 As Decimal


            '2022/1/29 MOD START 繰り返しなしの項目番号に繰り返し「〇」の入力があった場合エラーとする
            '9件目「繰り返し設定ができない項番に「繰り返し」設定が行われております。」追加
            '10件目「繰り返し設定しなければならない項番に「繰り返し」設定が行われていません。」追加
            Dim msg As String() = {"" _
                     , "{0}件目：{1}行目　「項目番号①」、「項目番号②」、「値③」、「データ範囲下限」、「データ範囲上限」は全て入力してください。" _
                     , "{0}件目：{1}行目　「項目番号①」に存在しない項番が入力されております。" _
                     , "{0}件目：{1}行目　「項目番号②」に存在しない項番が入力されております。" _
                     , "{0}件目：{1}行目　「値③」は正の整数5桁までで入力してください。" _
                     , "{0}件目：{1}行目　「データ範囲下限」は整数15桁、小数2桁までで入力してください。" _
                     , "{0}件目：{1}行目　「データ範囲上限」は整数15桁、小数2桁までで入力してください。" _
                     , "{0}件目：{1}行目　「データ範囲上限」は「データ範囲下限」より大きな値で入力してください。" _
                     , "{0}件目：{1}行目　「繰り返し」に「○」以外が入力されております。" _
                     , "{0}件目：{1}行目　繰り返し設定ができない項番に「繰り返し」設定が行われております。" _
                     , "{0}件目：{1}行目　繰り返し設定しなければならない項番（{2}）に「繰り返し」設定が行われていません。"
            }
            '2022/1/29 MOD END

            '調査票項目マスタ取得
            Dim dtItem As DataTable
            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, 調査年_引数用)
            End Using

            Dim cnt As Integer = 0

            For Each dc As Dictionary(Of String, String) In lstDc
                '1）項目番号①、項目番号②、値③、データ範囲下限、データ範囲上限については、全て入力されているか。
                If dc("項目番号１").ToString.Equals(String.Empty) _
                    OrElse dc("項目番号２").ToString.Equals(String.Empty) _
                    OrElse dc("値").ToString.Equals(String.Empty) _
                    OrElse dc("下限").ToString.Equals(String.Empty) _
                    OrElse dc("上限").ToString.Equals(String.Empty) Then
                    cnt = cnt + 1
                    details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                    ret = False
                    If cnt = max Then Return ret
                End If

                '2）項目番号①が存在するか。
                If Not dc("項目番号１").ToString.Equals(String.Empty) Then
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号１").ToString Select dr
                    If Not query.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '3）項目番号②が存在するか。
                If Not dc("項目番号２").ToString.Equals(String.Empty) Then
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号２").ToString Select dr
                    If Not query.Count > 0 Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '4）値③が5桁までの正数であるか。
                If Not dc("値").ToString.Equals(String.Empty) Then
                    If Long.TryParse(dc("値"), val1) Then
                        If Not (val1 > 0 And val1 < 10 ^ 5) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    Else
                        cnt = cnt + 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '5）データ範囲下限が正の整数15桁まで、少数2桁（半角数値で0.01から999999999999999.99までの入力）までであるか。
                If Not dc("下限").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("下限").ToString, "^[0-9]{1,15}(\.[0-9]{1,2})?$") Then
                        '---REV002 ADD START
                        '負の整数も許容する。
                        If Not Regex.IsMatch(dc("下限").ToString, "^-[0-9]{1,16}(\.[0-9]{1,2})?$") Then
                            '---REV002 ADD END
                            cnt = cnt + 1
                            details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                            '---REV002 ADD START
                        End If
                        '---REV002 ADD END
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("下限"), val) Then
                            '---REV002 MOD START
                            '負の整数を許容するので、絶対値にて判定
                            If Not (Math.Abs(val) >= 0.01D And Math.Abs(val) <= 999999999999999.99D) Then
                                'If Not (val >= 0.01D And val <= 999999999999999.99D) Then
                                '---REV002 MOD END
                                cnt = cnt + 1
                                details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If
                End If

                '6）データ範囲上限が正の整数15桁まで、少数2桁（半角数値で0.01から999999999999999.99までの入力）までであるか。
                If Not dc("上限").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("上限").ToString, "^[0-9]{1,15}(\.[0-9]{1,2})?$") Then
                        '---REV002 ADD START
                        '負の整数も許容する。
                        If Not Regex.IsMatch(dc("上限").ToString, "^-[0-9]{1,16}(\.[0-9]{1,2})?$") Then
                            '---REV002 ADD END
                            cnt = cnt + 1
                            details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                            '---REV002 ADD START
                        End If
                        '---REV002 ADD END
                    Else
                        Dim val As Decimal
                        If Decimal.TryParse(dc("上限"), val) Then
                            '---REV002 MOD START
                            '負の整数を許容するので、絶対値にて判定
                            If Not (Math.Abs(val) >= 0.01D And Math.Abs(val) <= 999999999999999.99D) Then
                                'If Not (val >= 0.01D And val <= 999999999999999.99D) Then
                                '---REV002 MOD END
                                cnt = cnt + 1
                                details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If
                End If

                '7）データ範囲上限 > データ範囲下限でがあるか。
                If Not (dc("上限").ToString.Equals(String.Empty) Or dc("下限").ToString.Equals(String.Empty)) Then
                    If Decimal.TryParse(dc("上限"), val2) And Decimal.TryParse(dc("下限"), val3) Then
                        If Not val2 > val3 Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    Else
                        cnt = cnt + 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If

                '8） 繰り返しに「〇」以外の入力があるか
                If Not (dc("繰り返し").ToString.Equals(String.Empty) Or dc("繰り返し").ToString = "○") Then

                    cnt = cnt + 1
                    details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                    ret = False
                    If cnt = max Then Return ret
                End If

                '2022/1/29 ADD START 繰り返しなしの項目番号に繰り返し「〇」の入力があった場合エラーとする
                '2022/03/17 REV start 繰り返しに「〇」があるのに、項目番号1、項目番号2、どちらにも可変項目が設定されていないときエラーとする
                '9）繰り返しがない項目番号に繰り返し「〇」の設定がされていないか
                If (dc("繰り返し").ToString = "○") Then

                    ''2022/03/17 REV start 繰り返しに「〇」があるのに、項目番号1、項目番号2、どちらにも可変項目が設定されていないときエラーとするためのフラグ
                    Dim koumoku1_flag As Boolean = False
                    Dim koumoku2_flag As Boolean = False

                    '項目番号１が繰り返し対象かチェック
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号１").ToString
                    If query.Any() Then
                        If (Integer.Parse(query(0)("可変区分").ToString) = 0) Then
                            'cnt = cnt + 1
                            'details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4), dc("項目番号１").ToString))
                            'ret = False
                            'If cnt = max Then Return ret
                            koumoku1_flag = True
                        End If
                    End If
                    '項目番号２が繰り返し対象かチェック
                    Dim query2 = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号２").ToString
                    If query2.Any() Then
                        If (Integer.Parse(query2(0)("可変区分").ToString) = 0) Then
                            'cnt = cnt + 1
                            'details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4), dc("項目番号２").ToString))
                            'ret = False
                            'If cnt = max Then Return ret
                            koumoku2_flag = True
                        End If
                    End If

                    '項目番号1、項目番号2、どちらも可変項目でないならエラー
                    If (koumoku1_flag And koumoku2_flag) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                End If
                '2022/1/29 ADD END

                '2022/3/10 ADD START 繰り返しありの項目番号に繰り返し「〇」の入力がない場合エラーとする
                '10）繰り返しがある項目番号に繰り返し「〇」の設定がされているか
                If (dc("繰り返し").ToString = "") Then
                    '項目番号１が繰り返し対象かチェック
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号１").ToString
                    If query.Any() Then
                        If (Integer.Parse(query(0)("可変区分").ToString) = 1) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4), dc("項目番号１").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                    '項目番号２が繰り返し対象かチェック
                    Dim query2 = From dr In dtItem Where dr("項目番号").ToString = dc("項目番号２").ToString
                    If query2.Any() Then
                        If (Integer.Parse(query2(0)("可変区分").ToString) = 1) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg(10), cnt.ToString.PadLeft(2), dc("連番").PadLeft(4), dc("項目番号２").ToString))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
                End If
                '2022/3/10 ADD END

            Next

            Return ret
        End Function
    End Class

    ''' <summary>
    ''' 労働時間整理ファイルクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RoudouFile
        ''' <summary>
        ''' 労働時間整理ファイルデータ取得
        ''' </summary>
        ''' REV009 ADD
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <param name="_chosaNen"></param>
        ''' <param name="_chosaKbn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetRoudouData(dt As DataTable, xlSheets As Excel.Sheets, comObject As ComObjectProcess, _chosaNen As String, _chosaKbn As String, excelName As String) As Dictionary(Of String, DAOOther.労働時間整理ファイル項目)
            Dim ret As New Dictionary(Of String, DAOOther.労働時間整理ファイル項目)

            'ウォッチ用
            'Dim 行位置_は_ As Integer
            'Dim 列位置_は_ As Integer
            'Dim aaaa As Integer

            Dim kouban As String
            Dim sname As String
            Dim record_num As Integer
            Dim column_num As Integer

            '労働時間整理ファイル毎のシートリストをファイル名から取得
            '↓MOD MS 2022/01/25
            'Dim sheets = ComConst.労働時間整理ファイル.シートリスト.リスト(excelName)
            '"）"の位置を取得
            Dim lastPosition As Integer = excelName.IndexOf("）", 0)
            'Excelファイルの拡張子を取得
            Dim ext = System.IO.Path.GetExtension(excelName)
            '比較するファイル名を形成
            Dim targetExcelName As String = excelName.Substring(0, lastPosition + 1) & ext
            Dim sheets = ComConst.労働時間整理ファイル.シートリスト.リスト(targetExcelName)
            '↑MOD MS 2022/01/25
            Dim xlSheet As Excel.Worksheet = Nothing

            For Each sheet In sheets
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

                    'シート保護確認
                    If xlSheet.ProtectContents Then
                        xlSheet.Unprotect()
                    End If

                    Dim rng As Excel.Range = Nothing
                    Try
                        Dim arrData(,) As Object
                        '↓MOD MS 2022/01/25
                        'rng = xlSheet.Range(ComConst.労働時間整理ファイル.シートデータ範囲.範囲(excelName)(sheet.ToString))
                        rng = xlSheet.Range(ComConst.労働時間整理ファイル.シートデータ範囲.範囲(targetExcelName)(sheet.ToString))
                        '↑MOD MS 2022/01/25

                        arrData = DirectCast(rng.Value, Object(,))

                        Dim sheetName = ""
                        Dim roudouSheetName = ""
                        If _chosaKbn <> ComConst.調査区分.営農類型別経営統計_個人 And _chosaKbn <> ComConst.調査区分.営農類型別経営統計_法人 Then
                            roudouSheetName = ComConst.労働時間整理ファイル.労働時間シート名.シート名(_chosaKbn)
                        End If

                        If sheet.ToString = roudouSheetName Then
                            '労働時間部分のシートのみ変換対象
                            If _chosaKbn = ComConst.調査区分.牛乳生産費統計_個別 Or _chosaKbn = ComConst.調査区分.子牛生産費統計_個別 Or
                           _chosaKbn = ComConst.調査区分.乳用雄育成牛生産費統計_個別 Or _chosaKbn = ComConst.調査区分.交雑種育成牛生産費統計_個別 Or
                           _chosaKbn = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 Or _chosaKbn = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 Or
                           _chosaKbn = ComConst.調査区分.交雑種肥育牛生産費統計_個別 Or _chosaKbn = ComConst.調査区分.肥育豚生産費統計_個別 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_牛乳生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_子牛生産費 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_肥育豚生産費 Then
                                '畜産の時は以下固定
                                sheetName = "【12】労働時間"
                            Else
                                '畜産以外はそのまま
                                sheetName = roudouSheetName
                            End If
                        Else
                            '労働時間部分のシートでなければ変換しない
                            sheetName = sheet.ToString
                        End If

                        Dim query = From dr In dt Where dr("シート名").ToString = sheetName Select dr

                        For Each dr As DataRow In query
                            kouban = dr("項目番号").ToString
                            sname = dr("シート名").ToString
                            '2022/03/16 ADD START 労働整理ファイル取込で更新しない項番はスキップ
                            record_num = CInt(dr("行位置").ToString)
                            column_num = CInt(dr("列位置").ToString)
                            '↓DEL MS 2022/03/09
                            'スキップが不要のため除去
                            '↓INS MS 2022/01/25
                            '不要項目はスキップ
                            'If kouban.Equals("Q12010501") OrElse kouban.Equals("Q12020201") OrElse kouban.Equals("Q12020301") OrElse kouban.Equals("Q12020401") Then
                            If _chosaKbn = ComConst.調査区分.牛乳生産費統計_個別 Or _chosaKbn = ComConst.調査区分.子牛生産費統計_個別 Or
                            _chosaKbn = ComConst.調査区分.乳用雄育成牛生産費統計_個別 Or _chosaKbn = ComConst.調査区分.交雑種育成牛生産費統計_個別 Or
                            _chosaKbn = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 Or _chosaKbn = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 Or
                            _chosaKbn = ComConst.調査区分.交雑種肥育牛生産費統計_個別 Or _chosaKbn = ComConst.調査区分.肥育豚生産費統計_個別 Then

                                If (record_num <= 7 Or record_num >= 110 Or column_num > 50 Or (record_num > 71 And record_num <= 73)) Then
                                    Continue For
                                End If


                            ElseIf _chosaKbn = ComConst.調査区分.経営分析調査_牛乳生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_子牛生産費 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 Or
                           _chosaKbn = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_肥育豚生産費 Then

                                If (record_num >= 32 Or column_num > 28) Then
                                    Continue For
                                End If

                            ElseIf _chosaKbn = ComConst.調査区分.米生産費統計_個別 Or _chosaKbn = ComConst.調査区分.小麦生産費統計_個別 Or
                                    _chosaKbn = ComConst.調査区分.二条大麦生産費統計_個別 Or _chosaKbn = ComConst.調査区分.六条大麦生産費統計_個別 Or
                                    _chosaKbn = ComConst.調査区分.はだか麦生産費統計_個別 Or _chosaKbn = ComConst.調査区分.そば生産費統計_個別 Or
                                    _chosaKbn = ComConst.調査区分.原料用かんしょ生産費統計_個別 Or _chosaKbn = ComConst.調査区分.原料用ばれいしょ生産費統計_個別 Or
                                    _chosaKbn = ComConst.調査区分.大豆生産費統計_個別 Or _chosaKbn = ComConst.調査区分.なたね生産費統計_個別 Or
                                    _chosaKbn = ComConst.調査区分.てんさい生産費統計_個別 Or _chosaKbn = ComConst.調査区分.さとうきび生産費統計_個別 Then

                                If (record_num >= 42 Or kouban.Equals(ComConst.労賃単価反映.労賃単価カラム(_chosaKbn)) Or record_num = 37 Or record_num = 38) Then
                                    Continue For
                                End If

                            ElseIf _chosaKbn = ComConst.調査区分.米生産費統計_組織法人 Or _chosaKbn = ComConst.調査区分.小麦生産費統計_組織法人 Or
                                   _chosaKbn = ComConst.調査区分.大豆生産費統計_組織法人 Or _chosaKbn = ComConst.調査区分.経営分析調査_二条大麦生産費 Or
                                   _chosaKbn = ComConst.調査区分.経営分析調査_六条大麦生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_はだか麦生産費 Or
                                   _chosaKbn = ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_そば生産費 Or
                                   _chosaKbn = ComConst.調査区分.経営分析調査_なたね生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_てんさい生産費 Or
                                   _chosaKbn = ComConst.調査区分.経営分析調査_さとうきび生産費 Then

                                If (record_num >= 46 Or kouban.Equals(ComConst.労賃単価反映.労賃単価カラム(_chosaKbn)) Or record_num = 40) Then
                                    Continue For
                                End If

                            ElseIf _chosaKbn = ComConst.調査区分.営農類型別経営統計_個人 Then

                                If kouban.Equals("Q10010701") Then
                                    Continue For
                                End If


                            End If

                            '↑INS MS 2022/01/25
                            '↑DEL MS 2022/03/09
                            '2022/03/16 ADD END 労働整理ファイル取込で更新しない項番はスキップ

                            If dr("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                                Dim item As New DAOOther.労働時間整理ファイル項目
                                With item

                                    'ウォッチ用
                                    'If ret.Count = 76 Then
                                    '    aaaa = 1
                                    'End If

                                    'ウォッチ用
                                    'If 行位置_は_ = 61 And 列位置_は_ = 12 Then
                                    '    aaaa = 2
                                    'End If

                                    .シート名 = dr("シート名").ToString
                                    '.行位置 = Integer.Parse(dr("行位置").ToString)
                                    '行位置_は_ = .行位置                               'ウォッチ用
                                    '↓MOD MS 2022/01/25
                                    .列位置 = Integer.Parse(dr("列位置").ToString)
                                    If _chosaKbn = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 Or
                                       _chosaKbn = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 Or _chosaKbn = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 Or
                                       _chosaKbn = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Then
                                        '★入力Excelファイルは行位置がずれているため、１を加算して補正
                                        .行位置 = Integer.Parse(dr("行位置").ToString) + 1
                                    Else
                                        .行位置 = Integer.Parse(dr("行位置").ToString)
                                    End If
                                    '↑MOD MS 2022/01/25
                                    '列位置_は_ = .列位置                                'ウォッチ用

                                    .型区分 = dr("型区分").ToString

                                    '営農個人　「10_労働」シート　表示行数特別対応
                                    If sheetName = "10_労働" And kouban = "Q10010701" Then
                                        .有効桁数 = 3
                                        .小数点以下桁数 = 0
                                        .値 = "100"
                                    Else
                                        .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                                        .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                                        '20220201 ADD START
                                        '行0列0のデータは飛ばす
                                        If String.IsNullOrEmpty(dr("行位置").ToString) OrElse dr("行位置").ToString = "0" OrElse
                                         String.IsNullOrEmpty(dr("列位置").ToString) OrElse dr("列位置").ToString = "0" Then
                                            Continue For
                                        End If
                                        '20220201 ADD END
                                        '↓INS MS 2022/01/25
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        '↑INS MS 2022/01/25
                                    End If
                                End With
                                ret.Add(dr("項目番号").ToString, item)
                            Else
                                Dim increment As Integer = Integer.Parse(dr("可変増量").ToString)
                                For i As Integer = 1 To Integer.Parse(dr("可変最大数").ToString)
                                    Dim item As New DAOOther.労働時間整理ファイル項目
                                    With item
                                        .シート名 = sheet.ToString '労働時間整理ファイルのシートを設定
                                        .行位置 = Integer.Parse(dr("行位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.縦, (i - 1) * increment, 0)
                                        '行位置_は_ = .行位置
                                        .列位置 = Integer.Parse(dr("列位置").ToString) + If(dr("可変方向").ToString = ComConst.可変方向.横, (i - 1) * increment, 0)
                                        '列位置_は_ = .列位置
                                        .値 = If(arrData(.行位置, .列位置) Is Nothing, Nothing, If(String.IsNullOrEmpty(arrData(.行位置, .列位置).ToString), Nothing, arrData(.行位置, .列位置).ToString))
                                        .型区分 = dr("型区分").ToString
                                    End With
                                    If Not String.IsNullOrEmpty(item.値) Then
                                        ret.Add(dr("項目番号").ToString & ComConst.ITEM_NO_DELIMITER & i.ToString, item)
                                    End If
                                Next
                            End If
                        Next
                    Finally
                        ReleaseComObject(comObject, rng)
                    End Try

                Finally
                    ReleaseComObject(comObject, xlSheet)
                End Try
            Next

            Return ret
        End Function

        ''' <summary>
        ''' 事務所番号変換（北海道対応）労働時間整理ファイル用
        ''' </summary>
        ''' <param name="sensasu"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvJimusyoNoRoudou(sensasu As String) As String
            Dim ret As String = String.Empty
            Dim ken As String = GetTodofuken(sensasu)
            Dim val As Integer

            If Integer.TryParse(ken, val) Then
                If val = 1 Then
                    val = 51
                    ret = val.ToString
                Else
                    ret = ken
                End If
            End If

            Return ret
        End Function

    End Class

    ''' <summary>
    ''' 事務所番号変換（北海道対応）調査票項目指定修正用
    ''' </summary>
    ''' <param name="sensasu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConvJimusyoNoKoumokuShitei(sensasu As String) As String
        Dim ret As String = String.Empty
        Dim ken As String = GetTodofuken(sensasu)
        Dim val As Integer

        If Integer.TryParse(ken, val) Then
            If val = 1 Then
                val = 51
                ret = val.ToString
            Else
                ret = ken
            End If
        End If

        Return ret
    End Function

    ''' <summary>
    ''' 所有牛情報更新時にクリアする平均的な飼養頭数の項番リストを取得します。(REV_019)
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function GetShiyotosuClearList() As List(Of String)
        Dim clearList = New List(Of String)
        Select Case CommonInfo.Chosakubun
            Case ComConst.調査区分.牛乳生産費統計_個別
                clearList.Add("Q02030801")
                clearList.Add("Q02030802")
                clearList.Add("Q02030803")
            Case ComConst.調査区分.子牛生産費統計_個別
                clearList.Add("Q02040802")
            Case ComConst.調査区分.去勢若齢肥育牛生産費統計_個別
                clearList.Add("Q02030801")
                clearList.Add("Q02030802")
            Case ComConst.調査区分.乳用雄育成牛生産費統計_個別,
                 ComConst.調査区分.交雑種育成牛生産費統計_個別,
                 ComConst.調査区分.乳用雄肥育牛生産費統計_個別,
                 ComConst.調査区分.交雑種肥育牛生産費統計_個別
                clearList.Add("Q02030801")
        End Select
        Return clearList
    End Function

#Region "制度受取金クラス"

    ''' <summary>
    ''' 制度受取金・積立金等項目クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Seidouketorikin

        ''' <summary>
        ''' シートデータ取得_取込（保存ボタン）用
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))

            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.制度受取金積立金等項目.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                Try

                    Dim arrData(,) As Object

                    rng1 = xlSheet.Range(ComConst.制度受取金積立金等項目.出力用ファイル名称.Col.First & ComConst.制度受取金積立金等項目.出力用ファイル名称.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.制度受取金積立金等項目.出力用ファイル名称.Col.First & ComConst.制度受取金積立金等項目.出力用ファイル名称.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = rng3.Row
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.制度受取金積立金等項目.出力用ファイル名称.Col.First & ComConst.制度受取金積立金等項目.出力用ファイル名称.Row.First & ":" _
                                            & ComConst.制度受取金積立金等項目.出力用ファイル名称.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.制度受取金積立金等項目.出力用ファイル名称.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' 制度受取金調査年コンボボックス設定
        ''' </summary>
        ''' <param name="heikinSyurui"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSeidouketorinenComboBox(Seidouketorinen As ComboBox)
            Seidouketorinen.DisplayMember = "Value"
            Seidouketorinen.ValueMember = "Key"
            Seidouketorinen.DataSource = ComUtil.Seidouketorikin.GetSeidouketorinen(False)
        End Sub

        ''' <summary>
        ''' 制度受取金調査年取得
        ''' </summary>
        ''' <param name="none"></param>
        ''' <returns></returns>
        Private Shared Function GetSeidouketorinen(Optional none As Boolean = False) As ArrayList
            Dim lst As New ArrayList()

            For Each Key In ComConst.制度受取金積立金等項目.制度受取金調査年.リスト.Keys
                lst.Add(New DictionaryEntry(Key, ComConst.制度受取金積立金等項目.制度受取金調査年.リスト(Key).調査年))
            Next

            Return lst
        End Function

        ''' <summary>調査区分</summary>
        Private Shared _Chosanen As String = Nothing

        ''' <summary>
        ''' 調査年取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Property Chosanen() As String

            Set(ByVal value As String)
                _Chosanen = value
            End Set
            Get
                Return _Chosanen
            End Get
        End Property

        ''' <summary>
        ''' シートデータ設定_Excel表示用
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.制度受取金積立金等項目.出力用ファイル名称.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.制度受取金積立金等項目.出力用ファイル名称.Col.First & ComConst.制度受取金積立金等項目.出力用ファイル名称.Row.First & ":" _
                                        & ComConst.制度受取金積立金等項目.出力用ファイル名称.Col.Last & dt.Rows.Count + ComConst.制度受取金積立金等項目.出力用ファイル名称.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.制度受取金積立金等項目.出力用ファイル名称.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                If protect Then
                    xlSheet.Protect()
                End If
            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub
    End Class
#End Region


    'REV 023 ADD START --------------
#Region "農業地域類型マスタ管理クラス"
    ''' <summary>
    ''' 農業地域類型マスタ管理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NogyoChikiMstMainte

        ''' <summary>
        ''' 農業地域類型マスタ管理コンボボックス設定
        ''' </summary>
        ''' <param name="nendo"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetMainteMastComboBox(nendo As ComboBox)
            nendo.DisplayMember = "Value"
            nendo.ValueMember = "Key"
            nendo.DataSource = ComUtil.NogyoChikiMstMainte.GetMainteMastaList(True)
        End Sub

        ''' <summary>
        ''' 農業地域類型マスタ取得
        ''' </summary>
        ''' <param name="none"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetMainteMastaList(Optional none As Boolean = False) As ArrayList
            Dim lst As New ArrayList()

            If none Then
                lst.Add(New DictionaryEntry(Nothing, Nothing))
            End If

            For Each Key In ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.リスト.Keys
                lst.Add(New DictionaryEntry(Key, ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.リスト(Key)))
            Next

            Return lst
        End Function

        ''' <summary>
        ''' 調査年に対応する農業地域類型マスタ調査年を返す
        ''' </summary>
        ''' <param name="chosanen"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetNogyoChikiMstYear(chosanen As String, chosaKubun As String) As String
            Dim ret As String
            '小麦類、なたね以外
            If Not (chosaKubun = ComConst.調査区分.小麦生産費統計_個別 Or
                chosaKubun = ComConst.調査区分.小麦生産費統計_組織法人 Or
                chosaKubun = ComConst.調査区分.二条大麦生産費統計_個別 Or
                chosaKubun = ComConst.調査区分.六条大麦生産費統計_個別 Or
                chosaKubun = ComConst.調査区分.はだか麦生産費統計_個別 Or
                chosaKubun = ComConst.調査区分.なたね生産費統計_個別) Then

                If ((CInt(chosanen)) <= 2021) Then
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2015
                ElseIf (2022 <= (CInt(chosanen)) And (CInt(chosanen)) <= 2026) Then
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2020
                ElseIf (2027 <= (CInt(chosanen)) And (CInt(chosanen)) <= 2031) Then
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2025
                Else
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2030
                End If
            Else '小麦類、なたね
                If ((CInt(chosanen)) <= 2022) Then
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2015
                ElseIf (2023 <= (CInt(chosanen)) And (CInt(chosanen)) <= 2027) Then
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2020
                ElseIf (2028 <= (CInt(chosanen)) And (CInt(chosanen)) <= 2032) Then
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2025
                Else
                    ret = ComConst.農業地域類型マスタ管理.農業地域類型マスタ調査年.農業地域類型マスタ2030
                End If
            End If
            Return ret
        End Function

        ''' <summary>
        ''' シートデータ取得
        ''' </summary>
        ''' <param name="xlSheets"></param>
        ''' <param name="comObject"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetData(xlSheets As Excel.Sheets, comObject As ComObjectProcess) As List(Of Dictionary(Of String, String))
            Dim ret As New List(Of Dictionary(Of String, String))
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.農業地域類型マスタ管理.ファイル情報.SheetName), Excel.Worksheet)

                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng1 As Excel.Range = Nothing
                Dim rng2 As Excel.Range = Nothing
                Dim rng3 As Excel.Range = Nothing
                Dim rngArr As Excel.Range = Nothing

                'シートデータ取得
                Try
                    Dim arrData(,) As Object
                    rng1 = xlSheet.Range(ComConst.農業地域類型マスタ管理.ファイル情報.Col.First & ComConst.農業地域類型マスタ管理.ファイル情報.Row.First)
                    If Not rng1.Value Is Nothing Then
                        Dim last As Integer

                        rng2 = xlSheet.Range(ComConst.農業地域類型マスタ管理.ファイル情報.Col.First & ComConst.農業地域類型マスタ管理.ファイル情報.Row.First + 1)
                        If Not rng2.Value Is Nothing Then
                            rng3 = rng1.End(Excel.XlDirection.xlDown)
                            last = If(rng3.Row <= ComConst.農業地域類型マスタ管理.ファイル情報.Row.Last, rng3.Row, ComConst.農業地域類型マスタ管理.ファイル情報.Row.Last)
                        Else
                            last = rng1.Row
                        End If

                        rngArr = xlSheet.Range(ComConst.農業地域類型マスタ管理.ファイル情報.Col.First & ComConst.農業地域類型マスタ管理.ファイル情報.Row.First & ":" _
                                               & ComConst.農業地域類型マスタ管理.ファイル情報.Col.Last & last)

                        arrData = DirectCast(rngArr.Formula, Object(,))

                        For i As Integer = LBound(arrData, 1) To UBound(arrData, 1)
                            Dim dc As New Dictionary(Of String, String)
                            For Each kv As KeyValuePair(Of Integer, String) In ComConst.農業地域類型マスタ管理.ファイル情報.Field
                                dc(kv.Value) = arrData(i, kv.Key).ToString
                            Next
                            ret.Add(dc)
                        Next
                    End If

                Finally
                    ReleaseComObject(comObject, rng1)
                    ReleaseComObject(comObject, rng2)
                    ReleaseComObject(comObject, rng3)
                    ReleaseComObject(comObject, rngArr)
                End Try

                'シート保護設定
                If protect Then
                    xlSheet.Protect()
                End If

            Finally
                ReleaseComObject(comObject, xlSheet)
            End Try

            Return ret
        End Function

        ''' <summary>
        ''' シートデータ設定
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="xlSheets"></param>
        ''' <param name="xlsProcess"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetSheetData(dt As DataTable, xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)
            Dim xlSheet As Excel.Worksheet = Nothing

            Try
                'シートの設定
                xlSheet = DirectCast(xlSheets.Item(ComConst.農業地域類型マスタ管理.ファイル情報.SheetName), Excel.Worksheet)

                'シート保護設定
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim rng As Excel.Range = Nothing
                Try
                    '明細一覧
                    Dim arrData(,) As Object

                    rng = xlSheet.Range(ComConst.農業地域類型マスタ管理.ファイル情報.Col.First & ComConst.農業地域類型マスタ管理.ファイル情報.Row.First & ":" _
                        & ComConst.農業地域類型マスタ管理.ファイル情報.Col.Last & dt.Rows.Count + ComConst.農業地域類型マスタ管理.ファイル情報.Row.First - 1)

                    arrData = DirectCast(rng.Formula, Object(,))

                    For i As Integer = 1 To dt.Rows.Count
                        For Each kv As KeyValuePair(Of Integer, String) In ComConst.農業地域類型マスタ管理.ファイル情報.Field
                            arrData(i, kv.Key) = dt.Rows(i - 1)(kv.Value).ToString
                        Next
                    Next

                    rng.Value = arrData
                    rng.Value = rng.Formula

                    SetAutoFit(xlsProcess, rng)
                Finally
                    ReleaseComObject(xlsProcess, rng)
                End Try

                'シート保護設定
                If protect Then
                    xlSheet.Protect()
                End If

            Finally
                ReleaseComObject(xlsProcess, xlSheet)
            End Try
        End Sub

        ''' <summary>
        ''' エラーチェック
        ''' </summary>
        ''' <param name="lstDc"></param>
        ''' <param name="details"></param>
        ''' <returns></returns>
        Public Shared Function CheckError(lstDc As List(Of Dictionary(Of String, String)), ByRef details As List(Of String)) As Boolean
            Dim ret As Boolean = True

            Const max As Integer = ComConst.ERR_MESSAGE_MAX

            Dim msg As String() = {"" _
                , "{0}件目：{1}行目　市区町村、旧市区町村、第１次分類、第２次分類は全て入力してください。" _
                , "{0}件目：{1}行目　都道府県、市区町村、旧市区町村の組み合わせがファイル上で重複しています。" _
                , "{0}件目：{1}行目　都道府県は半角数字２桁以内で入力してください。" _
                , "{0}件目：{1}行目　市区町村は半角数字３桁以内で入力してください。" _
                , "{0}件目：{1}行目　旧市区町村は半角数字２桁以内で入力してください。" _
                , "{0}件目：{1}行目　市区町村名は２０桁以内で入力してください。" _
                , "{0}件目：{1}行目　旧市区町村名は２０桁以内で入力してください。" _
                , "{0}件目：{1}行目　第１次分類は半角数字１桁以内で入力してください。" _
                , "{0}件目：{1}行目　第２次分類は半角数字１桁以内で入力してください。"
            }

            Dim row As Integer = 0
            Dim cnt As Integer = 0

            For Each dc As Dictionary(Of String, String) In lstDc
                row += 1

                '1）市区町村、旧市区町村、第１次分類、第２次分類について、全て入力されているか。
                If dc("市区町村").ToString.Equals(String.Empty) _
                    OrElse dc("旧市区町村").ToString.Equals(String.Empty) _
                    OrElse dc("第１次分類").ToString.Equals(String.Empty) _
                    OrElse dc("第２次分類").ToString.Equals(String.Empty) _
                    Then
                    cnt += 1
                    details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                    ret = False
                    If cnt = max Then Return ret
                End If

                '2）都道府県、市区町村、旧市区町村の組み合わせがファイル上で重複していないか。
                Dim query = From dct In lstDc Where dct("都道府県").PadLeft(2, "0"c) & dct("市区町村").PadLeft(3, "0"c) & dct("旧市区町村").PadLeft(2, "0"c) = dc("都道府県").PadLeft(2, "0"c) & dc("市区町村").PadLeft(3, "0"c) & dc("旧市区町村").PadLeft(2, "0"c)
                If query.Count > 1 Then
                    cnt += 1
                    details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                    ret = False
                    If cnt = max Then Return ret
                End If
                '3）都道府県が半角数字２桁以内であるか。
                If Not dc("都道府県").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("都道府県").ToString, "^[0-9]{1,2}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
                '4）市区町村が半角数字３桁以内であるか。
                If Not dc("市区町村").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("市区町村").ToString, "^[0-9]{1,3}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(4), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
                '5）旧市区町村が半角数字２桁以内であるか。
                If Not dc("旧市区町村").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("旧市区町村").ToString, "^[0-9]{1,2}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(5), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
                '6）市区町村名が入力されている場合、２０桁以内であるか。
                If Not dc("市区町村名").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("市区町村名").ToString, "^.{1,20}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(6), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
                '7）旧市区町村名が入力されている場合、２０桁以内であるか。
                If Not dc("旧市区町村名").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("旧市区町村名").ToString, "^.{1,20}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(7), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
                '8）第１次分類が半角数字１桁であるか。
                If Not dc("第１次分類").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("第１次分類").ToString, "^[0-9]{1}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(8), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
                '9）第２次分類が半角数字１桁であるか。
                If Not dc("第２次分類").ToString.Equals(String.Empty) Then
                    If Not Regex.IsMatch(dc("第２次分類").ToString, "^[0-9]{1}$") Then
                        cnt += 1
                        details.Add(String.Format(msg(9), cnt.ToString.PadLeft(2), row.ToString.PadLeft(4)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            Next

            Return ret
        End Function


    End Class
#End Region
    'REV 023 ADD END --------------
End Class