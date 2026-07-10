Imports Microsoft.Office.Interop

''' <summary>
''' 調査票出力
''' </summary>
''' <remarks></remarks>
'//*************************************************************************************************
'//  修正履歴
'// ------------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前      　　　　 |                修  正  内  容
'// -----------+------------+----------------------------+------------------------------------------
'//  REV_001   | 2020.11.16 |TSP)               　　　　 | フェーズ3 No10 対応
'//    
'//  REV_003   | 2021.12.20 |日本コンピュータシステム)   | 要件No.1-③対応
'//  REV_004   | 2021.12.29 |日本コンピュータシステム)   | 要件No.18対応
'//  REV_005   | 2023.10.30 |大興電子通信                | 変更要件No.6
'//*************************************************************************************************
'------------------------------------------------------------------------------------------
Public Class BRA9210R
    Inherits ExcelOutputMultipleBaseClass

    Private Const A4 As String = "A4"
    Private Const A3 As String = "A3"

    Private _shokuin As Boolean 'REV_001 ADD
    Private _yoshiSize As String
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <remarks></remarks>
     'REV-002 START----------------------
    '--- REV.001 MOD START
    Public Sub New(ByVal outDir As String, pshokuin As Boolean, _chosaNen As String, yoshiSize As String)
        'Public Sub New(ByVal outDir As String)
        '--- REV.001 MOD START
        MyBase.New(ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).tempFileName, True, False, ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(_chosaNen, CommonInfo.Chosakubun))).reportName, outDir, False)
        'REV-002 ENDT----------------------
        _shokuin = pshokuin   'REV_001 ADD 
        _yoshiSize = yoshiSize

        'REV_005↓調査票テンプレートは読取専用で開く
        IsReadOnly = True
        'REV_005↑
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(Of T As Class)(xlSheets As Excel.Sheets, unit As T)

        Dim pkey As DAOChosahyo.PrimaryKey = CType(DirectCast(unit, Object), DAOChosahyo.PrimaryKey)
        'REV-002 START-------
        Dim fileName As String = ComConst.調査票.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _
                                 & "_" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsm"
        'REV-002 END-------
        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        Dim dtItem As DataTable
        Dim dtChosahyo As Dictionary(Of String, DataTable)
        Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '調査票項目マスタ取得
            'REV-003 START（調査年追加）-------
            '20220202 MOD START 調査票出力時に制度受取金マスタの項目を取得するよう修正
            'dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, pkey.chosaNen, ComConst.数式区分.数式ではない)
            dtItem = DAOOther.GetChosahyoItemMasterSeidoUketori(db, CommonInfo.Chosakubun, pkey.chosaNen, ComConst.数式区分.数式ではない)
            '20220202 MOD END

            'REV-003 END-------

            '調査票テーブル取得
            dtChosahyo = DAOChosahyo.GetChosahyoTable(db, pkey)
        End Using

        '進捗加増
        Me.ProgressAddValue = 1

        '調査票項目取得
        dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

        '進捗加増
        Me.ProgressAddValue = 1

        '調査票シートデータ設定
        ComUtil.Chosahyo.SetSheetData(dcChosahyo, xlSheets, CType(Me, ComObjectProcess))

        '20220202 ADD START 調査票出力時に制度受取金マスタの項目を取得するよう修正
        'バージョン区分を判定する 新バージョンなら制度受取金セルロック設定
        If ComUtil.getVersionKubun(pkey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 Then
            Try
                '画面更新を無効にする
                xlApp.ScreenUpdating = False
                'イベント発生を無効にする
                xlApp.EnableEvents = False
                'アラート表示を無効にする
                xlApp.DisplayAlerts = False
                '全シート保護解除
                For Each xlSheet As Excel.Worksheet In xlSheets
                    xlSheet.Unprotect()
                Next

                '保護解除後に斜線とセルロックの設定
                Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)

            Finally
                '全シート保護
                For Each xlSheet As Excel.Worksheet In xlSheets
                    '保護設定後のオートフィルタ有効
                    xlSheet.Protect(AllowFiltering:=True)
                Next
                'アラート表示を有効にする
                xlApp.DisplayAlerts = True
                'イベント発生を有効にする
                xlApp.EnableEvents = True
                '画面更新を有効にする
                xlApp.ScreenUpdating = True
            End Try
        End If
        '20220202 ADD END

        '--- REV.001 ADD START
        '調査票表示・非表示設定
        If _shokuin = True Then
            setHyojiHihyoji(xlSheets)
        End If
        '--- REV.001 ADD END

        If _yoshiSize = A4 Then
            setYoshiSize(xlSheets, Excel.XlPaperSize.xlPaperA4)
        ElseIf _yoshiSize = A3 Then
            setYoshiSize(xlSheets, Excel.XlPaperSize.xlPaperA3)
        Else
            '従来通り（各テンプレの設定が有効）
        End If

        xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        '進捗加増
        Me.ProgressAddValue = 1
    End Sub

    '--- REV.001 ADD START
    Private Sub setHyojiHihyoji(xlSheets As Excel.Sheets)

        For Each xlSheet As Excel.Worksheet In xlSheets
            Try
                'シート保護確認
                Dim protect As Boolean = xlSheet.ProtectContents
                If protect Then
                    xlSheet.Unprotect()
                End If

                Dim hyoujiType As Integer = ComConst.調査票.非表示設定.なし

                If ComConst.調査票.シート非表示設定(CommonInfo.Chosakubun).ContainsKey(xlSheet.Name) Then
                    hyoujiType = ComConst.調査票.シート非表示設定(CommonInfo.Chosakubun).Item(xlSheet.Name)
                End If

                Select Case hyoujiType
                    Case ComConst.調査票.非表示設定.非表示
                        xlSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden
                    Case ComConst.調査票.非表示設定.折りたたみ
                        xlSheet.Outline.ShowLevels(0, 1)
                    Case ComConst.調査票.非表示設定.なし
                End Select

                If protect Then
                    '--- REV.004 MOD START
                    'xlSheet.Protect()
                    xlSheet.Protect(AllowFiltering:=True)
                    '--- REV.004 MOD START
                End If
            Finally
                ReleaseComObject(xlSheet)
            End Try
        Next
    End Sub
    '--- REV.001 ADD END

    Private Sub setYoshiSize(xlSheets As Excel.Sheets, paperSize As Excel.XlPaperSize)
        For Each xlSheet As Excel.Worksheet In xlSheets
            xlSheet.PageSetup.PaperSize = paperSize
            '20220311追記　用紙の倍率調整
            If paperSize = Excel.XlPaperSize.xlPaperA3 And CBool(xlSheet.PageSetup.Zoom) Then
                Dim zoom As Integer
                zoom = CInt(CLng(xlSheet.PageSetup.Zoom) * 1.4)
                xlSheet.PageSetup.Zoom = zoom

            End If
        Next
    End Sub
    '20220202 ADD START 調査票出力時に制度受取金マスタの項目を取得するよう修正
    ''' <summary>
    ''' 制度受取金項目セルロック設定
    ''' </summary>
    ''' <param name="dtItem"></param>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Private Sub SeidouketorikinCellLock(dtItem As DataTable, xlSheets As Excel.Sheets, dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目))
        Dim sheet As String = String.Empty
        Dim adress(151) As String
        Dim Seidokinkoubannum As Integer
        Dim Seidokinkouban As String

        If CommonInfo.Kubun2 = ComConst.区分２.畜産物生産費 Or CommonInfo.Chosakubun = ComConst.調査区分.原料用かんしょ生産費統計_個別 Then
            Exit Sub
        End If

        Select Case CommonInfo.Kubun2
            Case ComConst.区分２.営農類型別経営統計
                If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                    'adresstype = "A"
                    sheet = ComConst.制度受取金積立金等項目.制度受取金シート名(CommonInfo.Chosakubun)
                    For i = 1 To 152
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            If i <= 76 Then
                                Seidokinkoubannum = 9010110
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            Else
                                Seidokinkoubannum = 9010134
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            End If

                        End Using
                    Next i
                Else
                    'adresstype = "B"
                    sheet = ComConst.制度受取金積立金等項目.制度受取金シート名(CommonInfo.Chosakubun)
                    For i = 1 To 152
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            If i <= 76 Then
                                Seidokinkoubannum = 1010110
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q1" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            Else
                                Seidokinkoubannum = 1010134
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q1" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            End If
                        End Using
                    Next i
                End If

            Case ComConst.区分２.農産物生産費
                If CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 Then
                    'adresstype = "B"
                    sheet = ComConst.制度受取金積立金等項目.制度受取金シート名(CommonInfo.Chosakubun)
                    For i = 1 To 152
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            If i <= 76 Then
                                Seidokinkoubannum = 1110110
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            Else
                                Seidokinkoubannum = 1110134
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            End If
                        End Using
                    Next i

                ElseIf CommonInfo.Chosakubun = ComConst.調査区分.原料用ばれいしょ生産費統計_個別 Or
                        CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費 Then
                    'adresstype = "D"
                    sheet = ComConst.制度受取金積立金等項目.制度受取金シート名(CommonInfo.Chosakubun)
                    For i = 1 To 152
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            If i <= 76 Then
                                Seidokinkoubannum = 1060110
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            Else
                                Seidokinkoubannum = 1060134
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            End If
                        End Using
                    Next i

                ElseIf CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
                    'adresstype = "E"
                    sheet = ComConst.制度受取金積立金等項目.制度受取金シート名(CommonInfo.Chosakubun)
                    For i = 1 To 152
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            If i <= 76 Then
                                Seidokinkoubannum = 1130110
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            Else
                                Seidokinkoubannum = 1130134
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            End If
                        End Using
                    Next i
                Else
                    'adresstype = "C"
                    sheet = ComConst.制度受取金積立金等項目.制度受取金シート名(CommonInfo.Chosakubun)
                    For i = 1 To 152
                        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                            If i <= 76 Then
                                Seidokinkoubannum = 1070110
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            Else
                                Seidokinkoubannum = 1070134
                                Seidokinkoubannum = Seidokinkoubannum + i
                                Seidokinkouban = "Q0" & Seidokinkoubannum
                                Dim Seidokinkoubancheck = DAOOther.Seidokincheck2(db, CommonInfo.Chosakubun, ComUtil.Chosahyo.GetChosaNen(dcChosahyo), Seidokinkouban)
                                If Seidokinkoubancheck.Rows.Count <> 0 Then
                                    adress(i - 1) = Seidokinkouban
                                Else
                                    adress(i - 1) = "lock" & Seidokinkouban
                                End If
                            End If
                        End Using
                    Next i
                End If
        End Select

        Dim xlSheet As Excel.Worksheet = Nothing
        Try

            Dim cel As Excel.Range = Nothing
            Try

                For Each ad As String In adress
                    Dim adcheck As Boolean = ad.Contains("lock")


                    'adressの文字列にlockが含まれる（制度受取金等項番が指定されてない）場合、セルロック処理へ
                    If adcheck = False Then
                        Continue For
                    End If

                    '調査区分が営農法人の場合
                    ad = Strings.Right(ad, 9)
                    Dim sheetcheck As Integer = Array.IndexOf(ComConst.制度受取金積立金等項目.Seidouketorikin_sheetcheck, ad)
                    If sheet = "11_01_制度受取金・積立金" Or sheet = "11_02_制度受取金・積立金" Then
                        If sheetcheck = -1 Then
                            sheet = "11_01_制度受取金・積立金"
                        Else
                            sheet = "11_02_制度受取金・積立金"
                        End If
                    End If
                    sheetcheck = 0

                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)
                    cel = xlSheet.Cells

                    Dim query = From dr In dtItem Where dr("項目番号").ToString = ad Select dr
                    Dim rng As Excel.Range = Nothing

                    Try
                        If String.IsNullOrEmpty(query(0)("行位置").ToString) OrElse query(0)("行位置").ToString = "0" OrElse
                               String.IsNullOrEmpty(query(0)("列位置").ToString) OrElse query(0)("列位置").ToString = "0" Then
                            Continue For
                        End If
                        rng = DirectCast(cel.Item(Integer.Parse(query(0)("行位置").ToString), Integer.Parse(query(0)("列位置").ToString)), Excel.Range)

                        Dim area As Excel.Range = rng.MergeArea
                        Try
                            area.Locked = True
                        Finally
                            ReleaseComObject(area)
                        End Try
                    Finally
                        ReleaseComObject(rng)
                    End Try
                Next
            Finally
                ReleaseComObject(cel)
            End Try

        Finally
            ReleaseComObject(xlSheet)
        End Try
    End Sub
    '20220202 ADD END
End Class