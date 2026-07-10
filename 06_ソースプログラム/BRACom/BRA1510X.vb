Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms

''' <summary>
''' 電子調査票入力・修正（EXCEL）
''' </summary>
''' <remarks></remarks>
''' 
'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前            |                修  正  内  容
'// -----------+------------+-------------------------+--------------------------------------------
'//  REV_001   | 2020.12.17 |TSP)                     | フェーズ3 不具合対応No10修正
'//  REV_002   | 2021.11.29 |日本コンピュータシステム | 要件No1-③対応
'//  REV_003   | 2021.12.23 |日本コンピュータシステム | 要件10対応
'//  REV_004   | 2021.12.23 |日本コンピュータシステム | 要件11対応
'//  REV_005   | 2021.12.24 |日本コンピュータシステム | 要件19対応
'//  REV_006   | 2022.01.21 |日本コンピュータシステム | 要件No3対応
'//  REV_007   | 2023.04.19 |DAiKO                    | 変更要件No.3対応
'//  REV_008   | 2023.11.27 |DAiKO                    | 変更要件No.20対応
'//  REV_009   | 2025.09.10 |GCU                      | 要件5対応
'//            |            |                         |
'//*************************************************************************************************
Public Class BRA1510X
    Inherits ExcelInputBaseClass

    ''' <summary>編集モード種別</summary>
    Public Enum 編集モード種別
        新規 = 1
        修正
    End Enum

    ''' <summary>エラー時コメント改行文字数</summary>
    Private Const エラー時コメント改行文字数 As Integer = 30
    ''' <summary>最大エラー表示件数</summary>
    Private Const 最大エラー表示件数 As Integer = 1000

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton
    ''' <summary>基本エラーチェックボタン</summary>
    Private WithEvents btnBaseErrorCheck As CommandButton
    ''' <summary>追加エラーチェックボタン</summary>
    Private WithEvents btnAddErrorCheck As CommandButton
    ''' <summary>エラーチェック一覧ボタン</summary>
    Private WithEvents btnErrorCheckList As CommandButton
    ''' <summary>範囲エラーチェックボタン</summary>
    Private WithEvents btnRangeErrorCheck As CommandButton

    ''' <summary>編集モード</summary>
    Private _editMode As 編集モード種別

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>主キー</summary>
    Private _pKey As DAOChosahyo.PrimaryKey
    ''' <summary>拠点キー</summary>
    Private _kKey As DAOChosahyo.KotenKey

    ''' <summary>前回コメント情報</summary>
    Public _preCommentInfoList As List(Of CommentInfo)
    'BRA-002  Start--------------
    Public _preCommentInfoListHozonCheck As List(Of CommentInfo)
    'BRA-002  END--------------
    ''' <summary>エラーチェック実施情報</summary>
    Private _errCheckFlgDictionary As Dictionary(Of String, Boolean) 'REV003 Add

    ''' <summary>エラーチェック一覧情報</summary>
    Private _errCheckListDictionary As Dictionary(Of String, List(Of Dictionary(Of String, String))) 'REV003 Add

    ''' <summary>エラーチェック一覧出力ファイルパス</summary>
    Private _outFilePathOfErrCheckList As String  'REV003 Add

    ''' <summary>ロックファイルstream</summary>
    Private _lockFileStream As System.IO.FileStream 'REV004 Add

    ''' <summary>コピーテンプレートファイルパス</summary>
    Private _delfilepath As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="pFilePath"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="jimusho"></param>
    ''' <param name="kyoten"></param>
    ''' <param name="pKey"></param>
    ''' <remarks></remarks>

    'REV002 START --------
    '↓2022/01/26 引数の追加
    'Public Sub New(ByRef frm As ExcelInputBaseForm, editMode As 編集モード種別, chosaNen As String)
    Public Sub New(ByRef frm As ExcelInputBaseForm, editMode As 編集モード種別, chosaNen As String, pKey As DAOChosahyo.PrimaryKey, ByRef lockFileStream As System.IO.FileStream)
        '↑2022/01/26 引数の追加
        MyBase.New(frm, ComUtil.GetTemplateCopyFilePath(pKey.chosaNen, pKey.censusNo), True)
        'REV002 END --------
        Try

            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            _editMode = editMode
            '↓2022/01/26 pKeyの設定
            _pKey = pKey
            '↑2022/01/26 pKeyの設定
            _kKey = New DAOChosahyo.KotenKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center)
            _preCommentInfoList = New List(Of CommentInfo)
            _lockFileStream = lockFileStream  'REV004 Add
            'REV 002 start---------
            _preCommentInfoListHozonCheck = New List(Of CommentInfo)
            'REV 002 end-----------            'REV003 START --------
            _errCheckFlgDictionary = New Dictionary(Of String, Boolean)
            _errCheckFlgDictionary.Add(ComConst.審査論理種別.基本チェック, False)
            _errCheckFlgDictionary.Add(ComConst.審査論理種別.追加チェック, False)
            _errCheckFlgDictionary.Add(ComConst.審査論理種別.範囲チェック, False)
            _errCheckListDictionary = New Dictionary(Of String, List(Of Dictionary(Of String, String)))
            _errCheckListDictionary.Add(ComConst.審査論理種別.基本チェック, New List(Of Dictionary(Of String, String)))
            _errCheckListDictionary.Add(ComConst.審査論理種別.追加チェック, New List(Of Dictionary(Of String, String)))
            _errCheckListDictionary.Add(ComConst.審査論理種別.範囲チェック, New List(Of Dictionary(Of String, String)))
            _delfilepath = ComUtil.GetTemplateCopyFilePath(pKey.chosaNen, pKey.censusNo)
            'REV003 END --------

            'Excel画面を表示する
            Me.ShowExcel()

            '処理待ち画面を閉じる
            Me.CloseWaitForm()

            '2023/03/14 ADD START 問合せ9796対応
            'Excelをアクティブにする
            If (ComUtil.getVersionKubun(pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020) Then
                AppActivate("C_" + _pKey.chosaNen + "_" + _pKey.censusNo + "_" + ComUtil.GetChosakubunName(CommonInfo.Chosakubun) + "_電子調査票.xlsm")
                SendKeys.Send("{F2}")
                SendKeys.Send("{ESC}")
            End If
            '2023/03/14 ADD END 問合せ9796対応

        Catch ex As Exception
            'Excel画面を閉じる
            Me.CloseExcel()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="editMode"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef frm As ExcelInputBaseForm, editMode As 編集モード種別, pKey As DAOChosahyo.PrimaryKey, kKey As DAOChosahyo.KotenKey, ByRef lockFileStream As System.IO.FileStream)
        'REV002 START --------
        'MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), "C" & pKey.chosaNen & "_" & pKey.censusNo & "_" & ComConst.調査票.入力用ファイル名称(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubun(pKey.chosaNen, CommonInfo.Chosakubun)))), True)
        MyBase.New(frm, ComUtil.GetTemplateCopyFilePath(pKey.chosaNen, pKey.censusNo), True)
        'REV002 END --------
        Try

            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            _editMode = editMode
            _pKey = pKey
            _kKey = kKey
            _preCommentInfoList = New List(Of CommentInfo)
            _lockFileStream = lockFileStream  'REV004 Add
            _preCommentInfoListHozonCheck = New List(Of CommentInfo)
            'REV003 START --------
            _errCheckFlgDictionary = New Dictionary(Of String, Boolean)
            _errCheckFlgDictionary.Add(ComConst.審査論理種別.基本チェック, False)
            _errCheckFlgDictionary.Add(ComConst.審査論理種別.追加チェック, False)
            _errCheckFlgDictionary.Add(ComConst.審査論理種別.範囲チェック, False)
            _errCheckListDictionary = New Dictionary(Of String, List(Of Dictionary(Of String, String)))
            _errCheckListDictionary.Add(ComConst.審査論理種別.基本チェック, New List(Of Dictionary(Of String, String)))
            _errCheckListDictionary.Add(ComConst.審査論理種別.追加チェック, New List(Of Dictionary(Of String, String)))
            _errCheckListDictionary.Add(ComConst.審査論理種別.範囲チェック, New List(Of Dictionary(Of String, String)))
            _delfilepath = ComUtil.GetTemplateCopyFilePath(pKey.chosaNen, pKey.censusNo)
            'REV003 END --------

            'データを設定する
            Me.SetData()

            'Excel画面を表示する
            Me.ShowExcel()

            '処理待ち画面を閉じる
            Me.CloseWaitForm()
            '2023/03/14 ADD START 問合せ9796対応
            'Excelをアクティブにする
            If (ComUtil.getVersionKubun(pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020) Then
                AppActivate("C_" + _pKey.chosaNen + "_" + _pKey.censusNo + "_" + ComUtil.GetChosakubunName(CommonInfo.Chosakubun) + "_電子調査票.xlsm")
                SendKeys.Send("{F2}")
                SendKeys.Send("{ESC}")
            End If
            '2023/03/14 ADD END 問合せ9796対応
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
            btnBaseErrorCheck = ComUtilStrictOff.GetExcelBtnBaseErrorCheck(uf)
            btnAddErrorCheck = ComUtilStrictOff.GetExcelBtnAddErrorCheck(uf)

            'REV003 start
            If (ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020) Then
                btnErrorCheckList = ComUtilStrictOff.GetExcelBtnErrorCheckList(uf)
                btnRangeErrorCheck = ComUtilStrictOff.GetExcelBtnRangeErrorCheck(uf)
            End If
            'REV003 end 

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
    ''' データを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetData()
        Try
            xlApp.Interactive = False

            'Excel前処理
            Me.BeforeExcel()

            Dim dtItem As DataTable
            Dim dtChosahyo As Dictionary(Of String, DataTable)
            Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)



            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                'REV_006 MOD START -----------------------
                '調査票項目マスタ取得
                'dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _pKey.chosaNen, ComConst.数式区分.数式ではない)
                dtItem = DAOOther.GetChosahyoItemMasterSeidoUketori(db, CommonInfo.Chosakubun, _pKey.chosaNen, ComConst.数式区分.数式ではない)
                'REV_006 MOD END   -----------------------

                '調査票テーブル取得
                dtChosahyo = DAOChosahyo.GetChosahyoTable(db, _pKey)
            End Using

            '調査票項目取得
            dcChosahyo = ComUtil.Chosahyo.GetItem(dtItem, dtChosahyo)

            '調査票シートデータ設定
            ComUtil.Chosahyo.SetSheetData(dcChosahyo, xlSheets, CType(Me, ComObjectProcess))

            KoshinClick()

            'セルロック設定
            Me.CellLock(dtItem, xlSheets)
            '>>> 2022/01/27
            'Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)
            'バージョン区分を判定する 新バージョンなら制度受取金セルロック設定
            If ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 Then
                Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)
            End If
            '<<<

        Catch ex As Exception
            Throw
        Finally
            'Excel後処理
            Me.AfterExcel()

            xlApp.Interactive = True
        End Try
    End Sub

    ''' <summary>
    ''' セルロック設定
    ''' </summary>
    ''' <param name="dtItem"></param>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Private Sub CellLock(dtItem As DataTable, xlSheets As Excel.Sheets)
        Dim sheet As String = String.Empty
        Dim adress As String() = {}

        Select Case CommonInfo.Kubun2
            Case ComConst.区分２.営農類型別経営統計
                sheet = "00_表紙"
                adress = {ComConst.調査票.項目番号.調査年,
                          ComConst.調査票.項目番号.営農類型別経営統計.営農類型,
                          ComConst.調査票.項目番号.都道府県,
                          ComConst.調査票.項目番号.市区町村,
                          ComConst.調査票.項目番号.旧市区町村,
                          ComConst.調査票.項目番号.農業集落,
                          ComConst.調査票.項目番号.調査区,
                          ComConst.調査票.項目番号.客体番号}
            Case ComConst.区分２.農産物生産費
                sheet = "指標部入力"
                adress = {ComConst.調査票.項目番号.調査年,
                          ComConst.調査票.項目番号.都道府県,
                          ComConst.調査票.項目番号.市区町村,
                          ComConst.調査票.項目番号.旧市区町村,
                          ComConst.調査票.項目番号.農業集落,
                          ComConst.調査票.項目番号.調査区,
                          ComConst.調査票.項目番号.客体番号}
            Case ComConst.区分２.畜産物生産費
                sheet = "表紙"
                If CommonInfo.Chosakubun = ComConst.調査区分.乳用雄育成牛生産費統計_個別 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.交雑種育成牛生産費統計_個別 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.交雑種肥育牛生産費統計_個別 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 _
                    Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Then

                    adress = {ComConst.調査票.項目番号.調査年,
                              ComConst.調査票.項目番号.都道府県,
                              ComConst.調査票.項目番号.市区町村,
                              ComConst.調査票.項目番号.旧市区町村,
                              ComConst.調査票.項目番号.農業集落,
                              ComConst.調査票.項目番号.調査区,
                              ComConst.調査票.項目番号.客体番号}
                Else
                    adress = {ComConst.調査票.項目番号.調査年,
                              ComConst.調査票.項目番号.都道府県,
                              ComConst.調査票.項目番号.市区町村,
                              ComConst.調査票.項目番号.旧市区町村,
                              ComConst.調査票.項目番号.農業集落,
                              ComConst.調査票.項目番号.調査区,
                              ComConst.調査票.項目番号.客体番号}
                End If
        End Select

        Dim xlSheet As Excel.Worksheet = Nothing
        Try
            xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)

            Dim cel As Excel.Range = Nothing
            Try
                cel = xlSheet.Cells

                For Each ad As String In adress
                    Dim query = From dr In dtItem Where dr("項目番号").ToString = ad Select dr

                    Dim rng As Excel.Range = Nothing
                    Try
                        '>>> 2022/01/27 行0列0のデータは飛ばす
                        If String.IsNullOrEmpty(query(0)("行位置").ToString) OrElse query(0)("行位置").ToString = "0" OrElse
                           String.IsNullOrEmpty(query(0)("列位置").ToString) OrElse query(0)("列位置").ToString = "0" Then
                            Continue For
                        End If
                        '<<< 2022/01/27
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
            '>>>2022/01/27 位置変更①
            'xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)
            '<<<

            Dim cel As Excel.Range = Nothing
            Try
                '>>>2022/01/27 位置変更②
                'cel = xlSheet.Cells
                '<<<

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

                    '>>>2022/01/27 位置変更①②
                    xlSheet = DirectCast(xlSheets.Item(sheet), Excel.Worksheet)
                    cel = xlSheet.Cells
                    '<<<

                    Dim query = From dr In dtItem Where dr("項目番号").ToString = ad Select dr
                    Dim rng As Excel.Range = Nothing

                    Try
                        '2022/1/28 ADD START 行0列0のデータは飛ばす
                        If String.IsNullOrEmpty(query(0)("行位置").ToString) OrElse query(0)("行位置").ToString = "0" OrElse
                           String.IsNullOrEmpty(query(0)("列位置").ToString) OrElse query(0)("列位置").ToString = "0" Then
                            Continue For
                        End If
                        '2022/1/28 ADD END
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


    ''' <summary>
    ''' Excel画面を閉じる
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub CloseExcel()

        Try
            'ボタンの解放
            ReleaseComObject(btnSaveClose)
            ReleaseComObject(btnNoSaveClose)
            ReleaseComObject(btnBaseErrorCheck)
            ReleaseComObject(btnAddErrorCheck)

            'REV003 start
            If (ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020) Then
                ReleaseComObject(btnErrorCheckList)
                ReleaseComObject(btnRangeErrorCheck)
            End If
            'REV003 end 

            'Excel画面を閉じる
            MyBase.CloseExcel()

            'ロックファイル解放  'REV004 Add
            If Not _lockFileStream Is Nothing Then
                ComUtil.DeleteLockFile(_lockFileStream)
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 保存ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnSaveClose_Click() Handles btnSaveClose.Click
        Dim ProgressDialog As New ProgressDialog()

        Try
            'REV-005 START-------------------
            'バージョン区分を取得し、２以上の場合のみ計算処理を実行する                                     'INS 2022/02/03
            Dim versionKubun As String = ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun)     'INS 2022/02/03
            If versionKubun >= "2" Then                                                                     'INS 2022/02/03
                If ModuleKeisanClick() = False Then
                    If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_051, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                        Return
                    End If
                End If
            End If                                                                                          'INS 2022/02/03
            'REV-005 END-------------------


            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 4
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                Dim dtItem As DataTable
                Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)



                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '調査票項目マスタ取得
                    'dtItem = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _pKey.chosaNen)
                    dtItem = DAOOther.GetChosahyoItemMasterSeidoUketori(db, CommonInfo.Chosakubun, _pKey.chosaNen)

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    '調査票シートデータ取得
                    'REV-002 START-------------------
                    dcChosahyo = ComUtil.Chosahyo.GetSheetData(dtItem, xlSheets, CType(Me, ComObjectProcess), _pKey.chosaNen)
                    'REV-002 END-------------------
                    '>>>2022/01/27
                    Dim dtChosahyo As DataTable
                    dtChosahyo = DAOChosahyo.GetChosahyoList(db, _pKey.chosaNen, _kKey.kyoku, _kKey.jimusho, _kKey.kyoten, Nothing, Nothing, Nothing)
                    Dim chosaSoshiki As String = Nothing
                    Dim tantoMeisho As String = Nothing

                    For Each row As DataRow In dtChosahyo.Rows
                        If _pKey.censusNo = row(3).ToString Then
                            chosaSoshiki = row(5).ToString
                            tantoMeisho = row(6).ToString
                        End If
                    Next
                    '<<<2022/01/27
                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dcChosahyo, _editMode, details) Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        'エラーメッセージ
                        Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                        Exit Sub
                    End If

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    If _editMode = 編集モード種別.新規 Then
                        '主キー設定
                        _pKey = New DAOChosahyo.PrimaryKey(ComUtil.Chosahyo.GetChosaNen(dcChosahyo), ComUtil.Chosahyo.GetCensusNo(dcChosahyo))

                        '拠点キー設定

                        '調査票データ実査設置拠点取得
                        Dim koten As DAOChosahyo.KotenKey
                        koten = DAOChosahyo.GetChosahyoKoten(db, _pKey)

                        '実査設置拠点存在チェック
                        If koten Is Nothing Then
                            Try
                                db.BeginTrans()

                                '調査票データ追加
                                DAOChosahyo.InsertChosahyoTable(db, _pKey, _kKey, dcChosahyo, chosaSoshiki, tantoMeisho)

                                db.CommitTrans()
                            Catch ex As Exception
                                db.RollBackTrans()
                                Throw ex
                            End Try

                            '編集モード種別変更
                            _editMode = 編集モード種別.修正

                            'セルロック設定
                            Me.CellLock(dtItem, xlSheets)
                            '>>> 2022/01/27
                            'Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)
                            'バージョン区分を判定する 新バージョンなら制度受取金セルロック設定
                            If ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 Then
                                Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)
                            End If
                            '<<<
                        Else
                            If Not (CInt(koten.kyoku) = CInt(_kKey.kyoku) And CInt(koten.jimusho) = CInt(_kKey.jimusho) And CInt(koten.kyoten) = CInt(_kKey.kyoten)) Then
                                '進捗ダイアログを閉じる
                                ProgressDialog.endDispose()
                                'エラーメッセージ
                                Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_009, {_pKey.chosaNen, MasterDao.GetJimusyoName(koten.jimusho), MasterDao.GetCenterName(koten.jimusho, koten.kyoten), _pKey.censusNo}, MsgBoxStyle.OkOnly)
                                Exit Sub
                            Else
                                '確認メッセージ
                                If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_005, {_pKey.censusNo}, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                                    '個別結果表存在チェック
                                    Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(_pKey.chosaNen, _pKey.censusNo)
                                    Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(_kKey.kyoku, _kKey.jimusho, _kKey.kyoten)
                                    Dim bln As Boolean = DAOKobetsuKekkahyo.CheckExist(db, pKobet, kKobet)

                                    If bln Then
                                        '確認メッセージ
                                        If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_017, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If

                                    Try
                                        db.BeginTrans()

                                        If bln Then
                                            '個別結果表データ削除
                                            DAOKobetsuKekkahyo.DeleteTable(db, pKobet, kKobet)
                                        End If

                                        '調査票データ削除
                                        DAOChosahyo.DeleteChosahyoTable(db, _pKey, _kKey)
                                        '調査票データ追加
                                        DAOChosahyo.InsertChosahyoTable(db, _pKey, _kKey, dcChosahyo, chosaSoshiki, tantoMeisho)

                                        db.CommitTrans()
                                    Catch ex As Exception
                                        db.RollBackTrans()
                                        Throw ex
                                    End Try

                                    '編集モード種別変更
                                    _editMode = 編集モード種別.修正

                                    'セルロック設定
                                    Me.CellLock(dtItem, xlSheets)
                                    '>>> 2022/01/27
                                    'Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)
                                    'バージョン区分を判定する 新バージョンなら制度受取金セルロック設定
                                    If ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 Then
                                        Me.SeidouketorikinCellLock(dtItem, xlSheets, dcChosahyo)
                                    End If
                                    '<<<
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        '個別結果表存在チェック
                        Dim pKobet As DAOKobetsuKekkahyo.PrimaryKey = New DAOKobetsuKekkahyo.PrimaryKey(_pKey.chosaNen, _pKey.censusNo)
                        Dim kKobet As DAOKobetsuKekkahyo.KyotenKey = New DAOKobetsuKekkahyo.KyotenKey(_kKey.kyoku, _kKey.jimusho, _kKey.kyoten)
                        Dim bln As Boolean = DAOKobetsuKekkahyo.CheckExist(db, pKobet, kKobet)

                        If bln Then
                            '確認メッセージ
                            If ProgressDialog.ShowMsgBox(MessageID.MSG_Q_014, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If

                        Try
                            db.BeginTrans()

                            If bln Then
                                '個別結果表データ削除
                                DAOKobetsuKekkahyo.DeleteTable(db, pKobet, kKobet)
                            End If


                            '調査票データ削除
                            DAOChosahyo.DeleteChosahyoTable(db, _pKey, _kKey)

                            '調査票データ追加
                            DAOChosahyo.InsertChosahyoTable(db, _pKey, _kKey, dcChosahyo, chosaSoshiki, tantoMeisho)

                            db.CommitTrans()
                        Catch ex As Exception
                            db.RollBackTrans()
                            Throw ex
                        End Try
                    End If

                    '進捗を進める
                    ProgressDialog.AddValue = 1
                End Using

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                '完了メッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_001, MsgBoxStyle.OkOnly)
            End If
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally

            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

            'REV-005 START（Excel後処理の位置変更）-------------------
            'Excel後処理
            Me.AfterExcel()
            'REV-005 END-------------------
        End Try
    End Sub

    'REV-005 START-------------------
    ''' <summary>
    ''' 計算処理実行
    ''' </summary>
    ''' <returns></returns>
    Public Function ModuleKeisanClick() As Boolean
        Dim chosaKubunName As String
        Dim ret_val As Boolean = True

        chosaKubunName = ComUtil.GetChosakubunName(CommonInfo.Chosakubun)

        If chosaKubunName = "牛乳生産費統計（個別）" Or
            chosaKubunName = "子牛生産費統計（個別）" Or
            chosaKubunName = "乳用雄育成牛生産費統計（個別）" Or
            chosaKubunName = "交雑種育成牛生産費統計（個別）" Or
            chosaKubunName = "去勢若齢肥育牛生産費統計（個別）" Or
            chosaKubunName = "乳用雄肥育牛生産費統計（個別）" Or
            chosaKubunName = "交雑種肥育牛生産費統計（個別）" Or
            chosaKubunName = "経営分析調査_牛乳生産費" Or
            chosaKubunName = "経営分析調査_子牛生産費" Or
            chosaKubunName = "経営分析調査_乳用雄育成牛生産費" Or
            chosaKubunName = "経営分析調査_交雑種育成牛生産費" Or
            chosaKubunName = "経営分析調査_去勢若齢肥育牛生産費" Or
            chosaKubunName = "経営分析調査_乳用雄肥育牛生産費" Or
            chosaKubunName = "経営分析調査_交雑種肥育牛生産費" Then

            Try
                '20220202 MOD START 調査票の修正で開くファイルは頭にCがつく
                'xlApp.Run("'C" + chosaKubunName + "_電子調査票.xlsm'!Module1.計算_Click")
                '20220202 MOD END

                '20220225 MOD START 調査票の修正で開くファイルは頭にC_調査年_センサス番号がつく
                xlApp.Run("'C_" + _pKey.chosaNen + "_" + _pKey.censusNo + "_" + chosaKubunName + "_電子調査票.xlsm'!Module1.計算_Click")
                '20220225 MOD END

                Dim Result As Boolean

                '20220202 MOD START 調査票の修正で開くファイルは頭にCがつく
                'Result = DirectCast(xlApp.Run("'C" + chosaKubunName + "_電子調査票.xlsm'!Module1.ErrorCheck"), Boolean)
                '20220202 MOD END

                '20220225 MOD START 調査票の修正で開くファイルは頭にC_調査年_センサス番号がつく
                Result = DirectCast(xlApp.Run("'C_" + _pKey.chosaNen + "_" + _pKey.censusNo + "_" + chosaKubunName + "_電子調査票.xlsm'!Module1.ErrorCheck"), Boolean)
                '20220225 MOD END

                If Result = False Then
                    Throw New System.Exception("計算処理エラーがあります")
                End If

            Catch ex As Exception

                ret_val = False

            End Try

        End If
        Return ret_val
    End Function
    'REV-005 END-------------------

    Public Function KoshinClick() As Boolean

        Dim chosaKubunName As String
        Dim ret_val As Boolean = True

        chosaKubunName = ComUtil.GetChosakubunName(CommonInfo.Chosakubun)

        For Each temp As KeyValuePair(Of String, String) In ComConst.還元資料.帳票名
            ''営農類型以外の調査票
            'If chosaKubunName <> "営農類型別経営統計（個人）" And chosaKubunName <> "営農類型別経営統計（法人）" And chosaKubunName = temp.Value Then
            If chosaKubunName <> "営農類型別経営統計（法人）" And chosaKubunName = temp.Value Then
                Try
                    '↓REV_007 調査票の計算を実行する
                    If ComUtil.getVersionKubun(_pKey.chosaNen, CommonInfo.Chosakubun) = ComConst.バージョン区分.調査票項目2020 AndAlso
                       (chosaKubunName = "牛乳生産費統計（個別）" Or
                        chosaKubunName = "子牛生産費統計（個別）" Or
                        chosaKubunName = "乳用雄育成牛生産費統計（個別）" Or
                        chosaKubunName = "交雑種育成牛生産費統計（個別）" Or
                        chosaKubunName = "去勢若齢肥育牛生産費統計（個別）" Or
                        chosaKubunName = "乳用雄肥育牛生産費統計（個別）" Or
                        chosaKubunName = "交雑種肥育牛生産費統計（個別）") Then
                        xlApp.Run("'C_" + _pKey.chosaNen + "_" + _pKey.censusNo + "_" + chosaKubunName + "_電子調査票.xlsm'!Module1.計算_Click", False)
                        Thread.Sleep(5000)
                    End If
                    '↑REV_007
                    xlApp.Run("'C_" + _pKey.chosaNen + "_" + _pKey.censusNo + "_" + chosaKubunName + "_電子調査票.xlsm'!ThisWorkbook.Workbook_Open")
                Catch ex As Exception
                    ''何かしらの理由で実行できない場合の挙動は呼び出し先に任せる
                    ret_val = False
                End Try
                ''複数当てはまることはないはずなので終了
                Exit For
            End If
        Next

        Return ret_val
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="dcChosahyo"></param>
    ''' <param name="editMode"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目), editMode As 編集モード種別, ByRef details As List(Of String)) As Boolean
        Dim ret As Boolean = True

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String
        Dim cnt As Integer = 0

        '---REV 002 START ----------------
        Dim commentInfoList As New List(Of CommentInfo)

        Me.ClearCommnetHozonCheck()
        '---REV 002 END

        If editMode = 編集モード種別.新規 Then
            '1）操作しているユーザが専門調査員の場合、その専門調査員が扱えるセンサス番号かどうか。
            msg = "{0}件目：操作可能なセンサス番号ではありません。"
            If CommonInfo.SenmonChosain Then
                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    If Not DAOOther.CheckSenmonChosainKyakutaiExist(db, CommonInfo.UserId, ComUtil.Chosahyo.GetCensusNo(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End Using
            End If

            If Not ret Then Return ret

            '2）必須項目が入力されているか。
            msg = "{0}件目：{1}の入力がありません。"

            '調査年
            If String.IsNullOrEmpty(ComUtil.Chosahyo.GetChosaNen(dcChosahyo)) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "調査年"))
                ret = False
                If cnt = max Then Return ret
            End If

            Select Case CommonInfo.Kubun2
                Case ComConst.区分２.営農類型別経営統計
                    '営農類型
                    If String.IsNullOrEmpty(ComUtil.Chosahyo.GetEinouRuike(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "営農類型"))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                Case ComConst.区分２.農産物生産費
                    '対象品目
                    If String.IsNullOrEmpty(ComUtil.Chosahyo.GetTaishoHinmoku(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "対象品目"))
                        ret = False
                        If cnt = max Then Return ret
                    End If

                    '経営種類
                    If String.IsNullOrEmpty(ComUtil.Chosahyo.GetKeieiShurui(dcChosahyo)) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "経営種類"))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                Case ComConst.区分２.畜産物生産費
                    If CommonInfo.Chosakubun = ComConst.調査区分.乳用雄育成牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.交雑種育成牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.去勢若齢肥育牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.乳用雄肥育牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.交雑種肥育牛生産費統計_個別 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄育成牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種育成牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費 _
                        Or CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_交雑種肥育牛生産費 Then

                        '生産費区分
                        If String.IsNullOrEmpty(ComUtil.Chosahyo.GetSeisanhiKubun(dcChosahyo)) Then
                            cnt = cnt + 1
                            details.Add(String.Format(msg, cnt.ToString.PadLeft(2), "生産費区分"))
                            ret = False
                            If cnt = max Then Return ret
                        End If
                    End If
            End Select

            msg = "{0}件目：センサス番号の各項目について入力がない箇所があります。"
            If String.IsNullOrEmpty(ComUtil.Chosahyo.GetTodofuken(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetShikuchoson(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetKyuShikuchoson(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetNogyoShuraku(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetChosaku(dcChosahyo)) _
                OrElse String.IsNullOrEmpty(ComUtil.Chosahyo.GetKyakutaiNo(dcChosahyo)) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            If Not ret Then Return ret

            '3）農業経営統計調査システムにログインしている実査設置拠点を管轄している都道府県と、電子調査票の都道府県（センサス番号内）が一致しているか。
            msg = "{0}件目：自管轄の都道府県と、取込む電子調査票の都道府県が異なっています。"
            If Not CommonInfo.Jimusyo = ComUtil.Chosahyo.ConvJimusyoNo(dcChosahyo).PadLeft(2, "0"c) Then
                cnt = cnt + 1
                details.Add(String.Format(msg, cnt.ToString.PadLeft(2)))
                ret = False
                If cnt = max Then Return ret
            End If

            If Not ret Then Return ret
        End If

        '1, 4）電子調査票の各項目が、調査票項目マスタの型と一致しているか。
        msg = "{0}件目：シート名：{1}　{2}の型が一致しません。"
        For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In dcChosahyo
            If Not String.IsNullOrEmpty(kv.Value.値) Then
                If kv.Value.型区分 = ComConst.型区分.数値型 Then
                    Dim val As Decimal
                    'REV_001 MOD START----
                    If Not Decimal.TryParse(kv.Value.値, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, val) Then
                        'If Not Decimal.TryParse(kv.Value.値, val) Then
                        'REV001 MOD END---
                        cnt = cnt + 1
                        details.Add(String.Format(msg, cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))
                        ret = False

                        'REV-002 start----------------
                        commentInfoList.Add(New CommentInfo With {.SheetName = kv.Value.シート名.ToString,
                                                              .Row = kv.Value.行位置,
                                                              .Col = kv.Value.列位置,
                                                    .Comment = (String.Format(msg, cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key)),
                                                    .ErrChkType = enmErorCheckType.絶対エラー})
                        'REV-002 END----------------

                        If cnt = max Then Return ret
                        'REV_001 ADD START----
                    Else
                        '指数表記となっている場合、小数点以下に変換する。
                        If val <> 0 And val < 0.0001 Then
                            kv.Value.値 = val.ToString("0.#####")
                        End If
                        'REV001 ADD END---
                    End If
                End If
            End If
        Next

        '2, 5）電子調査票の各項目が、データベースの桁数に収まっているか。
        msg = "{0}件目：シート名：{1}　{2}の桁数がデータベースの桁数を超えています。"
        For Each kv As KeyValuePair(Of String, DAOChosahyo.調査票項目) In dcChosahyo
            If Not String.IsNullOrEmpty(kv.Value.値) Then
                If kv.Value.型区分 = ComConst.型区分.数値型 Then
                    Dim val As Decimal
                    If Decimal.TryParse(kv.Value.値, val) Then
                        If kv.Value.有効桁数 > 0 Then
                            Dim pattern As String
                            If kv.Value.小数点以下桁数 > 0 Then
                                pattern = "^-?[0-9]{1," & kv.Value.有効桁数 - kv.Value.小数点以下桁数 & "}(\.[0-9]{1," & kv.Value.小数点以下桁数 & "})?$"
                            Else
                                pattern = "^-?[0-9]{1," & kv.Value.有効桁数 & "}$"
                            End If
                            If Not Regex.IsMatch(kv.Value.値, pattern) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg, cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))

                                commentInfoList.Add(New CommentInfo With {.SheetName = kv.Value.シート名.ToString,
                                                              .Row = kv.Value.行位置,
                                                              .Col = kv.Value.列位置,
                                                    .Comment = (String.Format(msg, cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key)),
                                                    .ErrChkType = enmErorCheckType.絶対エラー})


                                '-------REV002 Start-----------------
                                'セルの背景色を変更
                                'xlInterior.Color = 1
                                '---REV002 END----------------

                                ret = False
                                If cnt = max Then Return ret
                            End If
                        End If
                    End If
                End If
            End If
        Next

        'REV 002 START
        SetCommentInfo(commentInfoList)

        '前回コメント情報にセットする
        _preCommentInfoListHozonCheck.Clear()
        _preCommentInfoListHozonCheck.AddRange(commentInfoList)



        Return ret
    End Function

    ''' <summary>
    ''' 戻るボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnNoSaveClose_Click() Handles btnNoSaveClose.Click
        Try

            'Excel画面を閉じる
            Me.CloseExcel()

            'コピーしたテンプレートファイルの削除
            Kill(_delfilepath)

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    ''' <summary>
    ''' 基本エラーチェックボタン
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub btnBaseErrorCheck_Click() Handles btnBaseErrorCheck.Click
        Dim dtCheckRonri As DataTable
        Dim progressDialog As New ProgressDialog()

        Try
            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_009, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    dtCheckRonri = DAOOther.GetChosahyoShinsaRonri(db, ComConst.エラーチェック種別.enm.基本)
                    If dtCheckRonri.Rows.Count = 0 Then
                        '審査論理データ無し
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_025, MsgBoxStyle.OkOnly)
                        setInfoToErrCheckList(ComConst.審査論理種別.基本チェック, New List(Of Dictionary(Of String, String))) 'REV003 Add
                        Return
                    End If
                End Using

                '進捗ダイアログを表示する
                Dim adjust As Integer = CType(Math.Ceiling(dtCheckRonri.Rows.Count * 0.05), Integer)
                progressDialog.Maximum = dtCheckRonri.Rows.Count + adjust
                progressDialog.Show(_formHwnd)

                Try
                    'Excel前処理
                    Me.BeforeExcel()

                    Dim errCntZ As Integer = 0
                    Dim errCntW As Integer = 0
                    Dim isErrMaxCnt As Boolean = False
                    '調査票のエラーチェックを行う
                    Me.CheckChosahyo(ComConst.審査論理種別.基本チェック, errCntZ, errCntW, isErrMaxCnt, progressDialog)

                    '進捗を進める
                    progressDialog.AddValue = adjust

                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()

                    'エラー件数が1000件に達した場合
                    If isErrMaxCnt Then
                        '完了メッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_039, MsgBoxStyle.OkOnly)
                    Else
                        '完了メッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_008, {errCntZ.ToString, errCntW.ToString}, MsgBoxStyle.OkOnly)
                    End If

                    'エラー一覧を表示する
                    'REV003 START --------
                    If errCntZ <> 0 Or errCntW <> 0 Then
                        showErrorCheckListDialog(ComConst.審査論理種別.基本チェック)
                    End If
                    'REV003 END ----------
                Finally
                    'Excel後処理
                    Me.AfterExcel()
                End Try
            End If

        Catch ex As ShinsaException
            '進捗ダイアログを閉じる
            progressDialog.endDispose()
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_055, {ex.ErrSign}, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            If Not progressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                progressDialog.endDispose()
                progressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' 追加エラーチェックボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnAddErrorCheck_Click() Handles btnAddErrorCheck.Click
        Dim dtCheckRonri As DataTable
        Dim progressDialog As New ProgressDialog()

        Try
            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_010, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    dtCheckRonri = DAOOther.GetChosahyoShinsaRonri(db, ComConst.エラーチェック種別.enm.追加)
                    If dtCheckRonri.Rows.Count = 0 Then
                        '審査論理データ無し
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_026, MsgBoxStyle.OkOnly)
                        setInfoToErrCheckList(ComConst.審査論理種別.追加チェック, New List(Of Dictionary(Of String, String))) 'REV003 Add
                        Return
                    End If
                End Using

                '進捗ダイアログを表示する
                Dim adjust As Integer = CType(Math.Ceiling(dtCheckRonri.Rows.Count * 0.05), Integer)
                progressDialog.Maximum = dtCheckRonri.Rows.Count + adjust
                progressDialog.Show(_formHwnd)

                Try
                    'Excel前処理
                    Me.BeforeExcel()

                    Dim errCntZ As Integer = 0
                    Dim errCntW As Integer = 0
                    Dim isErrMaxCnt As Boolean = False
                    '調査票のエラーチェックを行う
                    Me.CheckChosahyo(ComConst.審査論理種別.追加チェック, errCntZ, errCntW, isErrMaxCnt, progressDialog)

                    '進捗を進める
                    progressDialog.AddValue = adjust

                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()

                    'エラー件数が1000件に達した場合
                    If isErrMaxCnt Then
                        '完了メッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_040, MsgBoxStyle.OkOnly)
                    Else
                        '完了メッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_009, {errCntZ.ToString, errCntW.ToString}, MsgBoxStyle.OkOnly)
                    End If

                    'エラー一覧を表示する
                    'REV003 START --------
                    If errCntZ <> 0 Or errCntW <> 0 Then
                        showErrorCheckListDialog(ComConst.審査論理種別.追加チェック)
                    End If
                    'REV003 END ----------
                Finally
                    'Excel後処理
                    Me.AfterExcel()
                End Try
            End If

        Catch ex As ShinsaException
            '進捗ダイアログを閉じる
            progressDialog.endDispose()
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_055, {ex.ErrSign}, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            If Not progressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                progressDialog.endDispose()
                progressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' エラーチェック一覧ボタン
    ''' </summary>
    ''' <remarks>
    ''' REV003 Add
    ''' </remarks>
    Private Sub btnErrorCheckList_Click() Handles btnErrorCheckList.Click
        Dim ProgressDialog As New ProgressDialog()

        '問い合わせメッセージを表示→「いいえ」なら処理終了
        If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_052, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.No Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Return
        End If

        '各エラーチェック未実施の場合、メッセージを表示
        If Not _errCheckFlgDictionary(ComConst.審査論理種別.基本チェック) Then
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_104, MsgBoxStyle.OkOnly, MessageBoxDefaultButton.Button2)
        End If

        If Not _errCheckFlgDictionary(ComConst.審査論理種別.追加チェック) Then
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_113, MsgBoxStyle.OkOnly, MessageBoxDefaultButton.Button2)
        End If

        If Not _errCheckFlgDictionary(ComConst.審査論理種別.範囲チェック) Then
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_114, MsgBoxStyle.OkOnly, MessageBoxDefaultButton.Button2)
        End If

        '全エラーチェック未実施の場合、処理終了
        If Not (_errCheckFlgDictionary(ComConst.審査論理種別.基本チェック) Or
                _errCheckFlgDictionary(ComConst.審査論理種別.追加チェック) Or
                _errCheckFlgDictionary(ComConst.審査論理種別.範囲チェック)) Then
            Return
        End If

        '出力するエラーがない場合メッセージ表示→処理終了
        If _errCheckListDictionary(ComConst.審査論理種別.基本チェック).Count() = 0 _
            And _errCheckListDictionary(ComConst.審査論理種別.追加チェック).Count() = 0 _
            And _errCheckListDictionary(ComConst.審査論理種別.範囲チェック).Count() = 0 Then

            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_116, MsgBoxStyle.OkOnly, MessageBoxDefaultButton.Button2) = MsgBoxResult.Ok Then
                Return
            End If

        End If

        'REV_009
        Try
            Dim fileName As String
            If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
                fileName = ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Report.reportName & "_" _
                                    & CommonInfo.ChosakubunName & "_" & _pKey.censusNo & ".xlsx"
            Else
                fileName = ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Report.reportName & "_" _
                     & CommonInfo.ChosakubunName & "_" & _pKey.censusNo & "_" & _pKey.chosaNen & "_" & Now.ToString("yyyyMMddHHmm") & ".xlsx"
            End If

            'ファイルパス取得
            Dim tws As New ComUtil.fileDialogWithNewThread(IniFileInfo.ExcelOutPath, fileName, AddressOf setFileName)
            Dim t As New Thread(New ThreadStart(AddressOf tws.GetFilePath))
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
            t.Join()

            If _outFilePathOfErrCheckList.Equals(String.Empty) Then
                Exit Sub
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            '進捗ダイアログを表示する
            Dim adjust As Integer = 1
            ProgressDialog.Maximum = _errCheckListDictionary(ComConst.審査論理種別.基本チェック).Count _
                                + _errCheckListDictionary(ComConst.審査論理種別.追加チェック).Count _
                                + _errCheckListDictionary(ComConst.審査論理種別.範囲チェック).Count
            ProgressDialog.Show(_formHwnd)

            'エラーチェック一覧出力を行う
            Try
                Dim ret As ExcelOutputBaseClass.enmOutputReturn
                Using ExcelOutput = New BRA1510R(_pKey, _errCheckListDictionary, _outFilePathOfErrCheckList, ProgressDialog)
                    ret = ExcelOutput.Execute(MessageID.MSG_Q_004)
                End Using

                '進捗を進める
                ProgressDialog.AddValue = adjust

                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()

                If ret = ExcelOutputBaseClass.enmOutputReturn.OK Then
                    '完了メッセージ
                    Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_002, MsgBoxStyle.OkOnly)  'CHG:2022/01/27 NCS_TM 「_formHwnd,」追加、Msgが後ろに隠れてしまう対策
                End If
            Catch ex As ExcelOutputBaseClass.SaveAsException
                'エラーメッセージ
                Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_006, MsgBoxStyle.OkOnly)      'CHG:2022/01/27 NCS_TM 「_formHwnd,」追加、Msgが後ろに隠れてしまう対策
            Catch ex As Exception
                Throw
            End Try

        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)          'CHG:2022/01/27 NCS_TM 「_formHwnd,」追加、Msgが後ろに隠れてしまう対策
        Finally
            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

            ' ToDo 2021/12/15
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    ''' <summary>
    ''' ファイル選択後のコールバック
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <remarks>
    ''' REV003 Add
    ''' </remarks>
    Public Sub setFileName(filePath As String)
        _outFilePathOfErrCheckList = filePath
    End Sub

    ''' <summary>
    ''' 範囲エラーチェックボタン
    ''' </summary>
    ''' <remarks>
    ''' REV003 Add
    ''' </remarks>
    Private Sub btnRangeErrorCheck_Click() Handles btnRangeErrorCheck.Click

        Dim dtCheckRonri As DataTable
        Dim progressDialog As New ProgressDialog()


        Try
            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_050, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    dtCheckRonri = DAOOther.GetChosahyoShinsaRonriRange(db)
                    If dtCheckRonri.Rows.Count = 0 Then
                        '審査論理データ無し
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_102, MsgBoxStyle.OkOnly)
                        setInfoToErrCheckList(ComConst.審査論理種別.範囲チェック, New List(Of Dictionary(Of String, String))) 'REV003 Add
                        Return
                    End If
                End Using

                '進捗ダイアログを表示する
                Dim adjust As Integer = CType(Math.Ceiling(dtCheckRonri.Rows.Count * 0.05), Integer)
                progressDialog.Maximum = dtCheckRonri.Rows.Count + adjust
                progressDialog.Show(_formHwnd)

                Try
                    'Excel前処理
                    Me.BeforeExcel()


                    Dim errCntH As Integer = 0
                    Dim isErrMaxCnt As Boolean = False
                    '調査票の範囲エラーチェックを行う
                    Me.CheckChosahyo(ComConst.審査論理種別.範囲チェック, dtCheckRonri, errCntH, isErrMaxCnt, progressDialog)

                    '進捗を進める
                    progressDialog.AddValue = adjust

                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()

                    'エラー件数が1000件に達した場合
                    If isErrMaxCnt Then
                        '完了メッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_103, MsgBoxStyle.OkOnly)
                    Else
                        '完了メッセージ
                        Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_047, {errCntH.ToString}, MsgBoxStyle.OkOnly)
                    End If

                    'エラー一覧を表示する
                    'REV003 START --------
                    If errCntH <> 0 Then
                        showErrorCheckListDialog(ComConst.審査論理種別.範囲チェック)
                    End If
                    'REV003 END ----------
                Finally
                    'Excel後処理
                    Me.AfterExcel()
                End Try
            End If

        Catch ex As ShinsaException
            '進捗ダイアログを閉じる
            progressDialog.endDispose()
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_055, {ex.ErrSign}, MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_999, MsgBoxStyle.OkOnly)
        Finally
            If Not progressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                progressDialog.endDispose()
                progressDialog = Nothing
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try


    End Sub

    ''' <summary>
    ''' 調査票のエラーチェックを行う
    ''' </summary>
    ''' <param name="pShinsaRonriType"></param>
    ''' <param name="pErrCntZ">エラー件数</param>
    ''' <param name="pErrCntW">ワーニング件数</param>
    ''' <param name="pIsErrMaxCnt">最大エラー表示件数に達したかどうか</param>
    ''' <param name="progressDialog"></param>
    ''' <remarks></remarks>
    Private Sub CheckChosahyo(ByVal pShinsaRonriType As String, ByRef pErrCntZ As Integer, ByRef pErrCntW As Integer, ByRef pIsErrMaxCnt As Boolean, ByVal progressDialog As ProgressDialog)
        Dim dtChoItemMst As DataTable
        Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)
        Dim errCheckList As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String)) 'REV003 Add

        Try

            '前回エラー箇所のコメント及び背景色をクリアする
            Me.ClearCommnet()

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _pKey.chosaNen)

                '調査票シートデータ取得
                'REV-002 START-------------------
                dcChosahyo = ComUtil.Chosahyo.GetSheetData(dtChoItemMst, xlSheets, CType(Me, ComObjectProcess), _pKey.chosaNen)
                'REV-002 END-------------------
            End Using

            Dim targetDataList As New List(Of Shinsa.TargetData)
            For Each dc In dcChosahyo
                Dim meisaiNo As Integer
                If dc.Key.Contains(ComConst.ITEM_NO_DELIMITER) Then
                    meisaiNo = CInt(dc.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1))
                Else
                    meisaiNo = 0
                End If
                targetDataList.Add(New Shinsa.TargetData With {.ItemNo = dc.Key, .Value = dc.Value.値, .MeisaiNo = meisaiNo})
            Next

            If _editMode = 編集モード種別.新規 Then
                '主キー設定
                _pKey = New DAOChosahyo.PrimaryKey(ComUtil.Chosahyo.GetChosaNen(dcChosahyo), ComUtil.Chosahyo.GetCensusNo(dcChosahyo))
            End If

            Dim shinsa As Shinsa = New Shinsa(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei),
                                              CommonInfo.Chosakubun,
                                              ComConst.審査論理データ種別.調査票,
                                              pShinsaRonriType,
                                              _pKey.chosaNen,
                                              _kKey.jimusho,
                                              targetDataList,
                                              progressDialog,
                                              最大エラー表示件数)

            '審査実行
            Dim shinsaKekkaList As List(Of Shinsa.ShinsaKekka) = shinsa.Execute()

            Dim errCnt As Integer
            Dim commentInfoList As New List(Of CommentInfo)
            'エラー件数をセット
            pErrCntZ = (From n In shinsaKekkaList(0).ShinsaInfoList Where n.Result = False AndAlso n.ErrKubun = ComConst.エラーレベル.エラー).Count
            'ワーニング件数をセット
            pErrCntW = (From n In shinsaKekkaList(0).ShinsaInfoList Where n.Result = False AndAlso n.ErrKubun = ComConst.エラーレベル.ワーニング).Count
            pIsErrMaxCnt = False

            'REV003 START --------
            ' エラーチェック種別名を取得
            Dim errCheckTypeName As String
            Select Case pShinsaRonriType
                Case ComConst.審査論理種別.基本チェック
                    errCheckTypeName = ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.基本)
                Case ComConst.審査論理種別.追加チェック
                    errCheckTypeName = ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.追加)
                Case Else
                    errCheckTypeName = ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.基本)
            End Select
            'REV003 END ----------

            For Each shinsaInfo In (From n In shinsaKekkaList(0).ShinsaInfoList Where n.Result = False)
                '最大エラー表示件数に達した場合
                If errCnt >= 最大エラー表示件数 Then
                    pIsErrMaxCnt = True
                    Exit For
                End If

                Dim itemNoList As New List(Of String)
                '「エラーとなる条件(繰り返し項目増加したもの)」から『[』『]』で囲まれた文字列検索する
                Dim matchList As MatchCollection = Regex.Matches(shinsaInfo.ErrJokenConv, Shinsa.C_SearchKakko)
                For Each mat As Match In matchList
                    Dim itemNo As String = mat.Value.Trim("["c, "]"c)
                    '括弧なし項目番号の重複しないリストを作成する
                    If Not itemNoList.Contains(itemNo) Then
                        itemNoList.Add(itemNo)
                    End If
                Next

                'コメントリスト・エラーチェック一覧出力用データを作成
                Dim errSheetList As List(Of String) = New List(Of String) 'REV003 Add
                Dim changeableRowList As List(Of String) = New List(Of String) 'REV003 Add
                For Each itemNo In itemNoList
                    Dim meisaiNo As Integer = 0
                    Dim drChoItemMst As DataRow
                    If itemNo.Contains(ComConst.ITEM_NO_DELIMITER) Then
                        drChoItemMst = (From n In dtChoItemMst.AsEnumerable Where n("項目番号").ToString = itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(0) Select n).FirstOrDefault
                        meisaiNo = CType(itemNo.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1), Integer)
                    Else
                        drChoItemMst = (From n In dtChoItemMst.AsEnumerable Where n("項目番号").ToString = itemNo Select n).FirstOrDefault
                    End If
                    If drChoItemMst Is Nothing Then
                        Continue For
                    End If

                    errSheetList.Add(drChoItemMst("シート名").ToString) 'REV003 Add

                    Dim row As Integer
                    Dim Col As Integer
                    If drChoItemMst("可変区分").ToString = ComConst.可変区分.可変項目ではない Then
                        row = CInt(drChoItemMst("行位置"))
                        Col = CInt(drChoItemMst("列位置"))
                        changeableRowList.Add("") 'REV003 Add
                    Else
                        'count、sumの場合
                        If meisaiNo = 0 Then
                            row = CInt(drChoItemMst("行位置"))
                            Col = CInt(drChoItemMst("列位置"))
                        Else
                            row = CInt(drChoItemMst("行位置")) + If(drChoItemMst("可変方向").ToString = ComConst.可変方向.縦, (meisaiNo - 1) * CInt(drChoItemMst("可変増量")), 0)
                            Col = CInt(drChoItemMst("列位置")) + If(drChoItemMst("可変方向").ToString = ComConst.可変方向.横, (meisaiNo - 1) * CInt(drChoItemMst("可変増量")), 0)
                        End If
                        changeableRowList.Add(row.ToString()) 'REV003 Add
                    End If
                    commentInfoList.Add(New CommentInfo With {.SheetName = drChoItemMst("シート名").ToString,
                                                              .Row = row,
                                                              .Col = Col,
                                                              .Comment = "[" & Trim(shinsaInfo.ErrSign) & "]" & vbLf & AddStrCRLF(shinsaInfo.ErrNaiyo, エラー時コメント改行文字数),
                                                              .ErrChkType = If(shinsaInfo.ErrKubun = ComConst.エラーレベル.エラー, enmErorCheckType.絶対エラー, enmErorCheckType.警告エラー)})

                Next

                errCnt += 1

                'REV003 START --------
                ' エラー帳票出力用データを設定
                Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)

                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーシート名) _
                                           , String.Join(ComConst.調査票エラーチェックリスト一覧.delimiter, errSheetList.ToArray()))
                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.項目番号) _
                                           , String.Join(ComConst.調査票エラーチェックリスト一覧.delimiter, itemNoList.ToArray()))
                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーサイン) _
                                           , shinsaInfo.ErrSign)
                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラー内容) _
                                           , AddStrCRLF(shinsaInfo.ErrNaiyo, エラー時コメント改行文字数))
                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーチェック種別) _
                                           , errCheckTypeName & shinsaInfo.ErrKubun)
                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.可変行番号) _
                                           , If(changeableRowList.Count = changeableRowList.Where(Function(changeablerow) changeablerow.Equals("")).Count _
                                                , "" _
                                                , String.Join(ComConst.調査票エラーチェックリスト一覧.delimiter, changeableRowList.ToArray())))

                errCheckList.Add(errCheckDictionary)
                'REV003 END ----------
            Next

            Dim commentInfoListUniq As New List(Of CommentInfo)
            For Each commentInfo In commentInfoList
                '追加したコメントリストに「シート名、行位置、列位置」が同じデータが存在するかをチェック
                Dim commentList As List(Of CommentInfo) = (From n In commentInfoListUniq
                                                           Where n.SheetName = commentInfo.SheetName _
                                                           And n.Row = commentInfo.Row _
                                                           And n.Col = commentInfo.Col).ToList
                If commentList.Any Then
                    '存在する場合はコメント内容を連結する
                    commentList(0).Comment &= vbLf & commentInfo.Comment
                    '絶対エラーの場合、上書きする(絶対エラー優先)
                    If commentInfo.ErrChkType = enmErorCheckType.絶対エラー Then
                        commentList(0).ErrChkType = enmErorCheckType.絶対エラー
                    End If
                Else
                    commentInfoListUniq.Add(commentInfo)
                End If
            Next

            ' エラーチェック一覧用にコメントを保持する
            setInfoToErrCheckList(pShinsaRonriType, errCheckList) 'REV003 Add

            Try
                'コメントを追加する
                SetCommentInfo(commentInfoListUniq)
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Throw
            End Try

            '前回コメント情報にセットする
            _preCommentInfoList.Clear()
            _preCommentInfoList.AddRange(commentInfoListUniq)

        Catch ex As Exception
            Throw
        End Try
    End Sub


    ''' <summary>
    ''' 調査票の範囲エラーチェックを行う
    ''' </summary>
    ''' <param name="pShinsaRonriType"></param>
    ''' <param name="pErrCntH">範囲エラー件数</param>
    ''' <param name="pIsErrMaxCnt">最大エラー表示件数に達したかどうか</param>
    ''' <param name="progressDialog"></param>
    ''' <remarks></remarks>
    Private Sub CheckChosahyo(ByVal pShinsaRonriType As String, dtRonri As DataTable, ByRef pErrCntH As Integer, ByRef pIsErrMaxCnt As Boolean, ByVal progressDialog As ProgressDialog)
        Dim dtChoItemMst As DataTable
        Dim dcChosahyo As Dictionary(Of String, DAOChosahyo.調査票項目)
        Dim errCheckList As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String)) 'REV003 Add
        Dim ret As New List(Of Dictionary(Of String, String))
        Dim commentInfoList As New List(Of CommentInfo)

        Dim errCheckTypeName As String
        Dim koumoku_mojisuu As Int32 = 9

        errCheckTypeName = ComConst.エラーチェック種別.リスト(ComConst.エラーチェック種別.enm.範囲)

        Try

            '前回エラー箇所のコメント及び背景色をクリアする
            Me.ClearCommnet()

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _pKey.chosaNen)

                '調査票シートデータ取得
                'REV-002 START-------------------
                dcChosahyo = ComUtil.Chosahyo.GetSheetData(dtChoItemMst, xlSheets, CType(Me, ComObjectProcess), _pKey.chosaNen)
                'REV-002 END-------------------
            End Using


            Dim targetDataList As New List(Of Shinsa.TargetData)
            For Each dc In dcChosahyo
                Dim meisaiNo As Integer
                If dc.Key.Contains(ComConst.ITEM_NO_DELIMITER) Then
                    meisaiNo = CInt(dc.Key.Split(Char.Parse(ComConst.ITEM_NO_DELIMITER))(1))
                Else
                    meisaiNo = 0
                End If
                targetDataList.Add(New Shinsa.TargetData With {.ItemNo = dc.Key, .Value = dc.Value.値, .MeisaiNo = meisaiNo})
            Next




            For Each row As DataRow In dtRonri.Rows

                '最大エラー表示件数に達した場合
                If pErrCntH >= 最大エラー表示件数 Then
                    pIsErrMaxCnt = True
                    Exit For
                End If

                Dim val1 As Decimal
                Dim val2 As Decimal
                Dim val3 As Decimal

                Dim tanshu As Decimal

                '可変項目ならば
                If (row("繰り返し").ToString = "○") Then


                    ''2022/03/17 REV start 項目番号1、項目番号2が可変項目か判定するためのフラグ
                    '可変項目の時はfalse
                    Dim koumoku1_flag As Boolean = False
                    Dim koumoku2_flag As Boolean = False

                    '項目番号１が可変項目かチェック
                    Dim query = From dr In dtChoItemMst Where dr("項目番号").ToString = row("項目番号１").ToString
                    If query.Any() Then
                        If (Integer.Parse(query(0)("可変区分").ToString) = 0) Then

                            koumoku1_flag = True
                        End If
                    End If
                    '項目番号２が可変項目かチェック
                    Dim query2 = From dr In dtChoItemMst Where dr("項目番号").ToString = row("項目番号２").ToString
                    If query2.Any() Then
                        If (Integer.Parse(query2(0)("可変区分").ToString) = 0) Then

                            koumoku2_flag = True
                        End If
                    End If


                    '可変区分のデータ数分ループ
                    For Each targetdata In targetDataList
                        'どちらも可変項目の場合
                        If (koumoku1_flag = False And koumoku2_flag = False) Then
                            If (row("項目番号２").ToString = Left(targetdata.ItemNo, koumoku_mojisuu)) Then


                                If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).値, val1) _
                                AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).値, val2) _
                                 AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                    tanshu = val2 / val1 * val3

                                    Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                    Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                    If Not (low <= tanshu And tanshu <= up) Then
                                        Dim dc As New Dictionary(Of String, String)

                                        pErrCntH = pErrCntH + 1



                                        commentInfoList.Add(New CommentInfo With {.SheetName = dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名,
                                                                      .Row = dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).行位置,
                                                                      .Col = dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).列位置,
                                                                      .Comment = row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。",
                                                                      .ErrChkType = enmErorCheckType.範囲エラー})

                                        'REV003 START --------
                                        ' エラー帳票出力用データを設定
                                        Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)

                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.項目番号), row("項目番号２").ToString)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーサイン), "")
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.可変行番号), targetdata.MeisaiNo.ToString)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラー内容), AddStrCRLF(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーチェック種別), errCheckTypeName)
                                        errCheckList.Add(errCheckDictionary)
                                        'REV003 END ----------


                                    End If

                                End If
                            End If

                            '項目番号1のみが可変項目の場合
                        ElseIf (koumoku1_flag = False) Then


                            If (row("項目番号２").ToString = Left(targetdata.ItemNo, koumoku_mojisuu)) Then

                                '項目番号1と項目番号が同じtargetdataListの可変の数分ループ
                                Dim query3 = From dr In targetDataList Where Left(dr.ItemNo, koumoku_mojisuu) = row("項目番号１").ToString
                                For Each kahen1_list In query3

                                    If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString & ComConst.ITEM_NO_DELIMITER & kahen1_list.MeisaiNo).値, val1) _
                                    AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString).値, val2) _
                                    AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                        tanshu = val2 / val1 * val3

                                        Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                        Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                        If Not (low <= tanshu And tanshu <= up) Then
                                            Dim dc As New Dictionary(Of String, String)

                                            pErrCntH = pErrCntH + 1



                                            commentInfoList.Add(New CommentInfo With {.SheetName = dcChosahyo(row("項目番号２").ToString).シート名,
                                                                      .Row = dcChosahyo(row("項目番号２").ToString).行位置,
                                                                      .Col = dcChosahyo(row("項目番号２").ToString).列位置,
                                                                      .Comment = row("項目番号２").ToString & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。",
                                                                      .ErrChkType = enmErorCheckType.範囲エラー})

                                            'REV003 START --------
                                            ' エラー帳票出力用データを設定
                                            Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)

                                            errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString).シート名)
                                            errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.項目番号), row("項目番号２").ToString)
                                            errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーサイン), "")
                                            errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.可変行番号), If(CInt(dcChosahyo(row("項目番号２").ToString).可変範囲) > 0, row.ToString, ""))
                                            errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラー内容), AddStrCRLF(row("項目番号２").ToString & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                            errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーチェック種別), errCheckTypeName)
                                            errCheckList.Add(errCheckDictionary)
                                            'REV003 END ----------


                                        End If

                                    End If

                                Next
                            End If


                            '項目番号2のみが可変項目の場合
                        ElseIf (koumoku2_flag = False) Then
                            If (row("項目番号２").ToString = Left(targetdata.ItemNo, koumoku_mojisuu)) Then
                                If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString).値, val1) _
                                    AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).値, val2) _
                                    AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                                    tanshu = val2 / val1 * val3

                                    Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                                    Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                                    If Not (low <= tanshu And tanshu <= up) Then
                                        Dim dc As New Dictionary(Of String, String)

                                        pErrCntH = pErrCntH + 1



                                        commentInfoList.Add(New CommentInfo With {.SheetName = dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名,
                                                                      .Row = dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).行位置,
                                                                      .Col = dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).列位置,
                                                                      .Comment = row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。",
                                                                      .ErrChkType = enmErorCheckType.範囲エラー})

                                        'REV003 START --------
                                        ' エラー帳票出力用データを設定
                                        Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)

                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo).シート名)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.項目番号), row("項目番号２").ToString)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーサイン), "")
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.可変行番号), targetdata.MeisaiNo.ToString)
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラー内容), AddStrCRLF(row("項目番号２").ToString & ComConst.ITEM_NO_DELIMITER & targetdata.MeisaiNo & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                        errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーチェック種別), errCheckTypeName)
                                        errCheckList.Add(errCheckDictionary)
                                        'REV003 END ----------


                                    End If

                                End If

                            End If


                        End If



                    Next





                Else
                    If Decimal.TryParse(dcChosahyo(row("項目番号１").ToString).値, val1) _
                            AndAlso Decimal.TryParse(dcChosahyo(row("項目番号２").ToString).値, val2) _
                            AndAlso Decimal.TryParse(row("値").ToString, val3) Then

                        ' --- REV_001 MOD START
                        'If val1 > 0 And val2 > 0 Then   
                        If val1 <> 0 Then
                            ' --- REV_001 MOD END

                            tanshu = val2 / val1 * val3

                            Dim low As Decimal = Decimal.Parse(row("下限").ToString)
                            Dim up As Decimal = Decimal.Parse(row("上限").ToString)

                            If Not (low <= tanshu And tanshu <= up) Then
                                Dim dc As New Dictionary(Of String, String)

                                pErrCntH = pErrCntH + 1



                                commentInfoList.Add(New CommentInfo With {.SheetName = dcChosahyo(row("項目番号２").ToString).シート名,
                                                                      .Row = dcChosahyo(row("項目番号２").ToString).行位置,
                                                                      .Col = dcChosahyo(row("項目番号２").ToString).列位置,
                                                                      .Comment = row("項目番号２").ToString & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。",
                                                                      .ErrChkType = enmErorCheckType.範囲エラー})

                                'エラー帳票出力用データを設定
                                'REV003 START - -------

                                Dim errCheckDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)

                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーシート名), dcChosahyo(row("項目番号２").ToString).シート名)
                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.項目番号), row("項目番号２").ToString)
                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーサイン), "")
                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.可変行番号), If(CInt(dcChosahyo(row("項目番号２").ToString).可変範囲) > 0, row.ToString, ""))
                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラー内容), AddStrCRLF(row("項目番号２").ToString & ":" & row("チェック項目名").ToString & "が下限「" & row("下限").ToString & "」から上限「" & row("上限").ToString & "」の範囲ではありません。", エラー時コメント改行文字数))
                                errCheckDictionary.Add(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーチェック種別), errCheckTypeName)
                                errCheckList.Add(errCheckDictionary)
                                'REV003 END ----------


                            End If
                        End If

                    End If

                    '進捗を進める
                    progressDialog.AddValue = 1
                End If
            Next



            If _editMode = 編集モード種別.新規 Then
                '主キー設定
                _pKey = New DAOChosahyo.PrimaryKey(ComUtil.Chosahyo.GetChosaNen(dcChosahyo), ComUtil.Chosahyo.GetCensusNo(dcChosahyo))
            End If





            Dim commentInfoListUniq As New List(Of CommentInfo)
            For Each commentInfo In commentInfoList
                '追加したコメントリストに「シート名、行位置、列位置」が同じデータが存在するかをチェック
                Dim commentList As List(Of CommentInfo) = (From n In commentInfoListUniq
                                                           Where n.SheetName = commentInfo.SheetName _
                                                           And n.Row = commentInfo.Row _
                                                           And n.Col = commentInfo.Col).ToList
                If commentList.Any Then
                    '存在する場合はコメント内容を連結する
                    commentList(0).Comment &= vbLf & commentInfo.Comment
                    '絶対エラーの場合、上書きする(絶対エラー優先)
                    If commentInfo.ErrChkType = enmErorCheckType.絶対エラー Then
                        commentList(0).ErrChkType = enmErorCheckType.絶対エラー
                    End If
                Else
                    commentInfoListUniq.Add(commentInfo)
                End If
            Next

            ' エラーチェック一覧用にコメントを保持する
            setInfoToErrCheckList(pShinsaRonriType, errCheckList) 'REV003 Add

            Try
                'コメントを追加する
                SetCommentInfo(commentInfoListUniq)
            Catch ex As Exception
                'システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
                Throw
            End Try

            '前回コメント情報にセットする
            _preCommentInfoList.Clear()
            _preCommentInfoList.AddRange(commentInfoListUniq)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 前回エラー箇所のコメント及び背景色をクリアする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearCommnet()
        Try
            Dim xlSheet As Excel.Worksheet = Nothing
            Dim xlCells As Excel.Range = Nothing
            Dim xlTab As Excel.Tab = Nothing

            '重複しないシート名リストを取得
            Dim sheetNameList As List(Of String) = (From n In _preCommentInfoList Select n.SheetName Distinct).ToList
            For Each sheetName As String In sheetNameList
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                    xlCells = DirectCast(xlSheet.Cells, Excel.Range)
                    xlTab = DirectCast(xlSheet.Tab, Excel.Tab)

                    'シート見出しの色を初期化
                    xlTab.Color = False

                    'シート名が一致するコメント情報リストでループさせる
                    For Each CommentInfo In (From n In _preCommentInfoList Where n.SheetName = sheetName).ToList

                        Dim xlCell As Excel.Range = Nothing
                        Dim xlInterior As Excel.Interior = Nothing

                        Try
                            '2022/1/28 ADD START セル位置「0,0」のコメントを削除しないように修正 IF文追加
                            If (Not (CommentInfo.Row = 0 Or CommentInfo.Col = 0)) Then
                                '2022/1/28 ADD END
                                xlCell = DirectCast(xlCells(CommentInfo.Row, CommentInfo.Col), Excel.Range)
                                xlInterior = DirectCast(xlCell.Interior, Excel.Interior)

                                'コメントクリア
                                xlCell.ClearComments()
                                'セルの背景色を白に変更
                                xlInterior.Color = enmCellColor.White
                                '2022/1/28 ADD START セル位置「0,0」のコメントを削除しないように修正 IF文の対となるEndを追加
                            End If
                            '2022/1/28 ADD END

                        Finally
                            ReleaseComObject(xlInterior)
                            ReleaseComObject(xlCell)
                        End Try
                    Next
                Finally
                    ReleaseComObject(xlTab)
                    ReleaseComObject(xlCells)
                    ReleaseComObject(xlSheet)
                End Try
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 文字列に指定された文字数ごとに改行コードを挿入する
    ''' </summary>
    ''' <param name="myString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddStrCRLF(ByVal myString As String, ByVal myLength As Integer) As String
        Dim ret As System.Text.StringBuilder = Nothing

        Try

            ret = New System.Text.StringBuilder

            If myLength = 0 Then
                ret.Append(myString)
                Exit Try
            End If

            For i As Integer = 0 To myString.Length - 1
                If i <> 0 AndAlso (i Mod myLength) = 0 Then
                    ret.Append(vbLf)
                End If

                ret.Append(myString(i))
            Next

        Catch ex As Exception
            Throw
        End Try

        Return ret.ToString
    End Function

    ''' <summary>
    ''' エラーチェック一覧用データをセットする
    ''' </summary>
    ''' <remarks>
    ''' REV003 Add
    ''' </remarks>
    ''' <param name="pShinsaRonriType"></param>
    ''' <param name="errCheckList"></param>
    Private Sub setInfoToErrCheckList(ByVal pShinsaRonriType As String, ByVal errCheckList As List(Of Dictionary(Of String, String)))

        '対象エラーチェックの実施有無フラグをtrueに
        _errCheckFlgDictionary(pShinsaRonriType) = True

        '対象エラーチェックのデータがある場合、削除
        If _errCheckListDictionary.ContainsKey(pShinsaRonriType) Then
            _errCheckListDictionary.Remove(pShinsaRonriType)
        End If

        _errCheckListDictionary.Add(pShinsaRonriType, errCheckList)
    End Sub

    ''' <summary>
    ''' エラー一覧を表示
    ''' </summary>
    ''' <remarks>
    ''' REV003 Add
    ''' </remarks>
    Private Sub showErrorCheckListDialog(ByVal pShinsaRonriType As String)
        Dim errCnt As Integer = 0
        Dim errMsgList As List(Of String) = New List(Of String)

        ' エラーメッセージ作成
        For Each errCheckDictionary In _errCheckListDictionary(pShinsaRonriType)
            Dim sbErrMsg As New StringBuilder()
            sbErrMsg.AppendLine((errCnt + 1) & "件目：" _
                                & "(" & errCheckDictionary(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラーシート名)) & ")")
            sbErrMsg.AppendLine("[" & errCheckDictionary(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.項目番号)) & "]" _
                                & errCheckDictionary(ComConst.調査票エラーチェックリスト一覧.出力用ファイル名称.Field(ComConst.調査票エラーチェックリスト一覧.enm.エラー内容)))

            errMsgList.Add(sbErrMsg.AppendLine.ToString())
            errCnt += 1
            If errCnt = ComConst.ERR_MESSAGE_MAX Then
                Exit For
            End If
        Next

        Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_024, {String.Join(vbCrLf, errMsgList)})
    End Sub


    ''' <summary>
    ''' 前回エラー箇所のコメント及び背景色をクリアする(保存時のエラーチェック用)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearCommnetHozonCheck()
        Try
            Dim xlSheet As Excel.Worksheet = Nothing
            Dim xlCells As Excel.Range = Nothing
            Dim xlTab As Excel.Tab = Nothing

            '重複しないシート名リストを取得
            Dim sheetNameList As List(Of String) = (From n In _preCommentInfoListHozonCheck Select n.SheetName Distinct).ToList
            For Each sheetName As String In sheetNameList
                Try
                    xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                    xlCells = DirectCast(xlSheet.Cells, Excel.Range)
                    xlTab = DirectCast(xlSheet.Tab, Excel.Tab)

                    'シート見出しの色を初期化
                    xlTab.Color = False

                    'シート名が一致するコメント情報リストでループさせる
                    For Each CommentInfo In (From n In _preCommentInfoListHozonCheck Where n.SheetName = sheetName).ToList

                        Dim xlCell As Excel.Range = Nothing
                        Dim xlInterior As Excel.Interior = Nothing

                        Try
                            '2022/1/28 ADD START セル位置「0,0」のコメントを削除しないように修正 IF文追加
                            If (Not (CommentInfo.Row = 0 Or CommentInfo.Col = 0)) Then
                                '2022/1/28 ADD END
                                xlCell = DirectCast(xlCells(CommentInfo.Row, CommentInfo.Col), Excel.Range)
                                xlInterior = DirectCast(xlCell.Interior, Excel.Interior)

                                'コメントクリア
                                xlCell.ClearComments()
                                'セルの背景色を白に変更
                                xlInterior.Color = enmCellColor.White
                                '2022/1/28 ADD START セル位置「0,0」のコメントを削除しないように修正 IF文の対となるEndを追加
                            End If
                            '2022/1/28 ADD END

                        Finally
                            ReleaseComObject(xlInterior)
                            ReleaseComObject(xlCell)
                        End Try
                    Next
                Finally
                    ReleaseComObject(xlTab)
                    ReleaseComObject(xlCells)
                    ReleaseComObject(xlSheet)
                End Try
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class
