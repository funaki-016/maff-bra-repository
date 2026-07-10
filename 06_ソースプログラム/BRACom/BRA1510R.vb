Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Text

''' <summary>
''' 調査票エラーチェックリスト 出力
''' </summary>
''' 
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前            |                修  正  内  容
'// -----------+------------+-------------------------+--------------------------------------------
'//  REV_001   | 2020.12.17 |日本コンピュータシステム | 要件10 ファイル追加
'//  REV_002   | 2025.09.10 |GCU                      | 要件5 調査票エラーチェックリスト印刷範囲修正
'//            |            |                         |
'//*************************************************************************************************

Public Class BRA1510R
    Inherits ExcelOutputSingleBaseClass

    ''' <summary>主キー</summary>
    Private _pKey As DAOChosahyo.PrimaryKey
    ''' <summary>チェックリスト</summary>
    Private _checkList As Dictionary(Of String, List(Of Dictionary(Of String, String)))
    ''' <summary>進捗ダイアログ</summary>
    Private _progressDialog As ProgressDialog

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="errType"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="checkList"></param>
    ''' <param name="outPath"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pKey As DAOChosahyo.PrimaryKey, ByVal errCheckList As Dictionary(Of String, List(Of Dictionary(Of String, String))), ByVal outPath As String, progressDialog As ProgressDialog)
        MyBase.New(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Report.tempFileName, True, False, ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Report.reportName, outPath, True)

        _pKey = pKey
        _checkList = errCheckList
        _progressDialog = progressDialog
    End Sub

    Protected Overrides Sub ReportEdit(xlSheets As Sheets)
        'シートデータ設定
        Me.SetSheetData(_checkList, xlSheets)
    End Sub

    ''' <summary>
    ''' シートデータ設定
    ''' </summary>
    ''' <param name="lst"></param>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Private Sub SetSheetData(lst As Dictionary(Of String, List(Of Dictionary(Of String, String))), xlSheets As Excel.Sheets)
        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Report.SheetName), Excel.Worksheet)

            Dim rng As Excel.Range = Nothing
            Try
                '調査年欄
                rng = xlSheet.Range("B4")
                rng.Value = _pKey.chosaNen
                ReleaseComObject(rng)

                '調査区分欄
                rng = xlSheet.Range("D4")
                rng.Value = ComUtil.GetChosakubunName(CommonInfo.Chosakubun)
                ReleaseComObject(rng)

                '都道府県欄
                rng = xlSheet.Range("E7")
                rng.Value = ComUtil.GetTodofuken(_pKey.censusNo)
                ReleaseComObject(rng)

                '市区町村欄
                rng = xlSheet.Range("G7")
                rng.Value = ComUtil.GetShikuchoson(_pKey.censusNo)
                ReleaseComObject(rng)

                '旧市区町村欄
                rng = xlSheet.Range("J7")
                rng.Value = ComUtil.GetKyuShikuchoson(_pKey.censusNo)
                ReleaseComObject(rng)

                '農業集落
                rng = xlSheet.Range("M7")
                rng.Value = ComUtil.GetNogyoShuraku(_pKey.censusNo)
                ReleaseComObject(rng)

                '調査区
                rng = xlSheet.Range("O7")
                rng.Value = ComUtil.GetChosaku(_pKey.censusNo)
                ReleaseComObject(rng)

                '客体番号
                rng = xlSheet.Range("P7")
                rng.Value = ComUtil.GetKyakutaiNo(_pKey.censusNo)
                ReleaseComObject(rng)

                '明細一覧
                Dim errCheckList As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String))
                errCheckList.AddRange(lst(ComConst.審査論理種別.基本チェック))
                errCheckList.AddRange(lst(ComConst.審査論理種別.追加チェック))
                errCheckList.AddRange(lst(ComConst.審査論理種別.範囲チェック))

                Dim arrData(,) As Object
                rng = xlSheet.Range(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Col.First & ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Row.First & ":" _
                            & ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Col.Last & errCheckList.Count() + ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Row.First - 1)

                arrData = DirectCast(rng.Formula, Object(,))

                For i As Integer = 1 To errCheckList.Count()
                    _progressDialog.AddValue = 1
                    For Each kv As KeyValuePair(Of Integer, String) In ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field
                        arrData(i, kv.Key) = errCheckList(i - 1)(kv.Value).ToString
                    Next
                Next

                rng.Value = arrData
                rng.Value = rng.Formula
                'REV_002
                Dim startcol As String = ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Col.First
                Dim endcol As String = ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Col.Last
                Dim startrow As Integer = comconst.調査票エラーチェックリスト一覧.出力用ファイル名称.row.First

                Dim lastdatarow As Integer = startrow + errchecklist.count - 1

                Dim lastcell As excel.range = DirectCast(xlsheet.cells.specialcells(excel.xlcelltype.xlcelltypelastcell), excel.range)
                Dim templatelastrow As Integer = lastcell.row
                releasecomobject(lastcell)

                If templatelastrow > lastdatarow Then
                    Dim delfrom As Integer = lastdatarow + 1
                    Dim delto As Integer = templatelastrow
                    xlSheet.range(startcol & delfrom.tostring(), endcol & delto.tostring()).entirerow.delete()

                End If
                '行高自動調整
                AdjustErrContentHeight(xlSheet, startrow, lastdatarow)

            Finally
                ReleaseComObject(rng)
            End Try
        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub

    ''' <summary>
    ''' エラーメッセージ欄（O～R列）の内容に応じて行の高さを自動調整する。
    ''' </summary>
    ''' <param name="ws">対象となる Excel ワークシート。</param>
    ''' <param name="startRow">処理を開始する行番号。</param>
    ''' <param name="endRow">処理を終了する行番号。</param>
    ''' <remarks>
    ''' O～R 列を一時的に結合し、列幅を合算した状態で AutoFit を実行することで、
    ''' 折り返し表示（WrapText）を考慮した正しい行高を算出する。
    ''' <para>
    ''' 処理後は列結合状態および元の列幅を必ず復元する。
    ''' </para>
    ''' </remarks>
    Private Sub AdjustErrContentHeight(ws As Excel.Worksheet, startRow As Integer, endRow As Integer)
        If endRow < startRow Then Exit Sub

        Dim colO As Excel.Range = Nothing
        Dim colP As Excel.Range = Nothing
        Dim colQ As Excel.Range = Nothing
        Dim colR As Excel.Range = Nothing

        Try
            colO = DirectCast(ws.Columns("O"), Excel.Range)
            colP = DirectCast(ws.Columns("P"), Excel.Range)
            colQ = DirectCast(ws.Columns("Q"), Excel.Range)
            colR = DirectCast(ws.Columns("R"), Excel.Range)

            ' ColumnWidth は Object → Double に明示変換（Strict On）
            Dim wO As Double = CDbl(colO.ColumnWidth)
            Dim wP As Double = CDbl(colP.ColumnWidth)
            Dim wQ As Double = CDbl(colQ.ColumnWidth)
            Dim wR As Double = CDbl(colR.ColumnWidth)

            For r As Integer = startRow To endRow
                Dim rr As Excel.Range = Nothing
                Dim lt As Excel.Range = Nothing

                Try
                    rr = ws.Range("O" & r.ToString(), "R" & r.ToString())

                    rr.UnMerge()
                    rr.Merge()

                    lt = ws.Range("O" & r.ToString())
                    Dim text As String = ""
                    If lt.Value IsNot Nothing Then text = lt.Value.ToString()
                    colO.ColumnWidth = wO + wP + wQ + wR

                    rr.UnMerge()
                    rr.WrapText = True
                    rr.ShrinkToFit = False
                    rr.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                    rr.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

                    lt.Value = text

                    rr.Rows.RowHeight = ws.StandardHeight
                    rr.Rows.AutoFit()

                    rr.Merge()
                    lt.Value = text

                Finally
                    ' 列幅は必ず戻す（途中例外でも崩れないように）
                    If colO IsNot Nothing Then colO.ColumnWidth = wO
                    If colP IsNot Nothing Then colP.ColumnWidth = wP
                    If colQ IsNot Nothing Then colQ.ColumnWidth = wQ
                    If colR IsNot Nothing Then colR.ColumnWidth = wR

                    ReleaseComObject(lt)
                    ReleaseComObject(rr)
                End Try
            Next

        Finally
            ReleaseComObject(colR)
            ReleaseComObject(colQ)
            ReleaseComObject(colP)
            ReleaseComObject(colO)
        End Try
    End Sub

End Class
