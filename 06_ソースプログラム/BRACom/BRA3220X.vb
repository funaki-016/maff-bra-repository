Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Text.RegularExpressions

''' <summary>
''' 個別結果表入力・修正（EXCEL）
''' </summary>
''' <remarks></remarks>
''' 

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2021.12.07 |日本コンピュータシステム | 要件No1-③対応
'//  REV_002   | 2022.10.11 |Daiko                    | 要件No1 バージョン区分追加
'//  REV_003   | 2022.11.21 |Daiko                    | 要件No1 制度受取金・積立金等項目名称条件処理追加
'//
'//*************************************************************************************************

Public Class BRA3220X
    Inherits ExcelInputBaseClass

    ''' <summary>保存ボタン</summary>
    Private WithEvents btnSaveClose As CommandButton
    ''' <summary>戻るボタン</summary>
    Private WithEvents btnNoSaveClose As CommandButton
    ''' <summary>再計算ボタン</summary>
    Private WithEvents btnCalculate As CommandButton

    ''' <summary>Excelユーザーフォームハンドル</summary>
    Private _formHwnd As Win32WindowWrapper

    ''' <summary>主キー</summary>
    Private _pKey As DAOKobetsuKekkahyo.PrimaryKey
    ''' <summary>拠点キー</summary>
    Private _kKey As DAOKobetsuKekkahyo.KyotenKey
    ''' <summary>欠測値補完</summary>
    Private _kessokuchiHokan As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <param name="pKey"></param>
    ''' <param name="kKey"></param>
    ''' <param name="kessokuchiHokan"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef frm As ExcelInputBaseForm, pKey As DAOKobetsuKekkahyo.PrimaryKey, kKey As DAOKobetsuKekkahyo.KyotenKey, kessokuchiHokan As String)
        ' REV_002↓
        'MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.個別結果表.入力用ファイル名称(CommonInfo.Chosakubun)), True)
        MyBase.New(frm, System.IO.Path.Combine(IniFileInfo.ExcelReportPath(), ComConst.個別結果表.入力用ファイル名称(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(pKey.chosaNen, CommonInfo.Chosakubun)))), True)
        ' REV_002↑
        Try
            _pKey = pKey
            _kKey = kKey
            _kessokuchiHokan = kessokuchiHokan

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
            btnCalculate = ComUtilStrictOff.GetExcelBtnCalculate(uf)

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
            Dim dtKobetsu As Dictionary(Of String, DataTable)
            Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
            Dim nousanFlg As Boolean = False
            Dim seidoUketoriChosaNen As String = Nothing

            'REV_003 START---------------
            '農産物生産費かつR4体系の場合、制度受取金・積立金等項目フラグ、調査年を設定する
            If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 And ComConst.バージョン区分.結果表等項目2022 = ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun) Then
                seidoUketoriChosaNen = _pKey.chosaNen
                nousanFlg = True
            End If
            'REV_003 END---------------

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '個別結果表項目マスタ取得
                ' REV_002↓
                'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), False, seidoUketoriChosaNen)
                ' REV_002 REV_003↑

                '個別結果表テーブル取得
                dtKobetsu = DAOKobetsuKekkahyo.GetTable(db, _pKey, _kessokuchiHokan)
            End Using

            '個別結果表項目取得
            dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtItem, dtKobetsu, nousanFlg) 'REV_003

            '個別結果表シートデータ設定
            ComUtil.KobetsuKekkahyo.SetSheetData(dcKobetsu, xlSheets, CType(Me, ComObjectProcess))
        Catch ex As Exception
            Throw
        Finally
            'Excel後処理
            Me.AfterExcel()

            xlApp.Interactive = True
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
            ReleaseComObject(btnCalculate)

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
        Dim ProgressDialog As New ProgressDialog()

        Try
            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_001, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                ProgressDialog.Maximum = 4
                ProgressDialog.Show(_formHwnd)

                'Excel前処理
                Me.BeforeExcel()

                Dim dtItem As DataTable
                Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '個別結果表項目マスタ取得
                    ' REV_002↓
                    'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
                    dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun))
                    ' REV_002↑

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    '個別結果表シートデータ取得
                    dcKobetsu = Me.GetSheetData(dtItem, xlSheets)

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    'エラーチェック
                    Dim details As New List(Of String)
                    If Not Me.CheckError(dcKobetsu, details) Then
                        '進捗ダイアログを閉じる
                        ProgressDialog.endDispose()
                        'エラーメッセージ
                        Message.ShowMsgForm(_formHwnd, MessageID.MSG_E_010, {String.Join(vbCrLf, details)})
                        Exit Sub
                    End If

                    '進捗を進める
                    ProgressDialog.AddValue = 1

                    Try
                        db.BeginTrans()

                        '個別結果表データ削除
                        DAOKobetsuKekkahyo.DeleteTable(db, _pKey, _kKey, _kessokuchiHokan)

                        '個別結果表データ追加
                        DAOKobetsuKekkahyo.InsertTable(db, _pKey, _kKey, dcKobetsu, _kessokuchiHokan)

                        db.CommitTrans()

                    Catch ex As Exception
                        db.RollBackTrans()
                        Throw ex
                    End Try

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
            'Excel後処理
            Me.AfterExcel()

            If Not ProgressDialog Is Nothing Then
                '進捗ダイアログを閉じる
                ProgressDialog.endDispose()
                ProgressDialog = Nothing
            End If

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
    ''' 再計算ボタン
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnCalculate_Click() Handles btnCalculate.Click
        Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
        Dim dtCreateRonri As DataTable
        Dim progressDialog As New ProgressDialog()
        Dim taisyakuKubun As String = Nothing

        Try

            '確認メッセージ
            If Message.ShowMsgBox(_formHwnd, MessageID.MSG_Q_021, MsgBoxStyle.YesNo, MessageBoxDefaultButton.Button2) = MsgBoxResult.Yes Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                '進捗ダイアログを表示する
                progressDialog.Show(_formHwnd)

                Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                    '個別結果表シートデータ取得(個別結果表項目マスタ取得時は裏項番含めない)
                    ' REV_002↓
                    'dcKobetsu = Me.GetSheetData(DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun), xlSheets, True)
                    dcKobetsu = Me.GetSheetData(DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun)), xlSheets, True)
                    ' REV_002↑
                    If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                        taisyakuKubun = dcKobetsu(ComConst.個別結果表.貸借対照表(ComConst.調査区分.営農類型別経営統計_個人)).値
                        '個別結果表作成論理＿営農個人取得(再計算する論理のみ取得する)
                        ' REV_002↓
                        'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, True)
                        dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonriEinouKobetsu(db, taisyakuKubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), True)
                        ' REV_002↑
                    Else
                        '個別結果表作成論理取得(再計算する論理のみ取得する)
                        ' REV_002↓
                        'dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, True)
                        dtCreateRonri = DAOOther.GetKobetsuKekkahyoSakuseiRonri(db, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), True)
                        ' REV_002↑
                    End If
                End Using
                If dtCreateRonri.Rows.Count = 0 Then
                    '個別結果表作成論理データ無し
                    Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_017, MsgBoxStyle.OkOnly)
                    Return
                End If

                progressDialog.Maximum = dtCreateRonri.Rows.Count

                Try

                    'Excel前処理
                    Me.BeforeExcel()

                    '再計算処理を行う
                    Me.Recalculate(dcKobetsu, dtCreateRonri, progressDialog)

                    '進捗ダイアログを閉じる
                    progressDialog.endDispose()

                    '完了メッセージ
                    Message.ShowMsgBox(_formHwnd, MessageID.MSG_I_014, MsgBoxStyle.OkOnly)
                Finally
                    'Excel後処理
                    Me.AfterExcel()
                End Try
            End If

        Catch ex As CreateKobetsuException
            'システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Err, ex)
            Message.ShowMsgBox(_formHwnd, MessageID.MSG_E_037, {ex.ItemNo}, MsgBoxStyle.OkOnly)

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
    ''' 個別結果表シートデータ取得
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="xlSheets"></param>
    ''' <param name="pIsCnvDec">値を数値変換するかどうか</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSheetData(dt As DataTable, xlSheets As Excel.Sheets, Optional isCnvDec As Boolean = False) As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
        Dim ret As New Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)

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

                    Dim page As Excel.PageSetup = xlSheet.PageSetup
                    Try
                        rng = xlSheet.Range(page.PrintArea)
                    Finally
                        ReleaseComObject(page)
                    End Try

                    arrData = DirectCast(rng.Value, Object(,))

                    Dim query = From dr In dt Where dr("シート名").ToString = sheet.ToString Select dr

                    For Each dr As DataRow In query
                        Dim item As New DAOKobetsuKekkahyo.個別結果表項目
                        With item
                            .シート名 = dr("シート名").ToString
                            .行位置 = Integer.Parse(dr("行位置").ToString)
                            .列位置 = Integer.Parse(dr("列位置").ToString)
                            .値 = ComUtil.KobetsuKekkahyo.GetData(arrData(.行位置, .列位置), dr("型区分").ToString, isCnvDec)
                            .型区分 = dr("型区分").ToString
                            .有効桁数 = Integer.Parse(dr("有効桁数").ToString)
                            .小数点以下桁数 = Integer.Parse(dr("小数点以下桁数").ToString)
                            .表示単位 = dr("表示単位").ToString
                        End With
                        ret.Add(dr("項目番号").ToString, item)
                    Next
                Finally
                    ReleaseComObject(rng)
                End Try

            Finally
                ReleaseComObject(xlSheet)
            End Try
        Next

        Return ret
    End Function

    ''' <summary>
    ''' エラーチェック
    ''' </summary>
    ''' <param name="dcKobetsu"></param>
    ''' <param name="details"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckError(dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目), ByRef details As List(Of String)) As Boolean
        Dim ret As Boolean = True

        Const max As Integer = ComConst.ERR_MESSAGE_MAX

        Dim msg As String() = {"" _
                             , "{0}件目：シート名：{1}　{2}の型が一致しません。" _
                             , "{0}件目：シート名：{1}　{2}の桁数がデータベースの桁数を超えています。" _
                             , "{0}件目：シート名：{1}　{2}の桁数が個別結果表作成論理の表示単位桁数を超えています。" _
        }

        Dim cnt As Integer = 0

        '1）個別結果表の各項目が、個別結果表項目マスタの型と一致しているか。
        For Each kv As KeyValuePair(Of String, DAOKobetsuKekkahyo.個別結果表項目) In dcKobetsu
            If Not String.IsNullOrEmpty(kv.Value.値) Then
                If kv.Value.型区分 = ComConst.型区分.数値型 Then
                    Dim val As Decimal
                    If Not Decimal.TryParse(kv.Value.値, val) Then
                        cnt = cnt + 1
                        details.Add(String.Format(msg(1), cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))
                        ret = False
                        If cnt = max Then Return ret
                    End If
                End If
            End If
        Next

        '2）個別結果表の各項目が、データベースの桁数に収まっているか。
        '3）個別結果表作成論理に表示単位が設定されている個別結果表項目が、表示単位の桁数に収まっているか。
        For Each kv As KeyValuePair(Of String, DAOKobetsuKekkahyo.個別結果表項目) In dcKobetsu
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
                            If Not Regex.IsMatch(val.ToString, pattern) Then
                                cnt = cnt + 1
                                details.Add(String.Format(msg(2), cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))
                                ret = False
                                If cnt = max Then Return ret
                            Else
                                If kv.Value.小数点以下桁数 > 0 Then
                                    If Not String.IsNullOrEmpty(kv.Value.表示単位) Then
                                        Dim digit As Integer
                                        Dim unit As Decimal = Decimal.Parse(kv.Value.表示単位)
                                        If ComConst.個別結果表作成論理.表示単位.リスト.ContainsKey(unit) Then
                                            digit = kv.Value.表示単位.TrimEnd("0"c).Substring(kv.Value.表示単位.TrimEnd("0"c).IndexOf("."c) + 1).Length
                                        End If
                                        If digit > 0 Then
                                            pattern = "^-?[0-9]{1," & kv.Value.有効桁数 - kv.Value.小数点以下桁数 & "}(\.[0-9]{1," & digit & "})?$"
                                        Else
                                            pattern = "^-?[0-9]{1," & kv.Value.有効桁数 - kv.Value.小数点以下桁数 & "}$"
                                        End If
                                        If Not Regex.IsMatch(val.ToString, pattern) Then
                                            cnt = cnt + 1
                                            details.Add(String.Format(msg(3), cnt.ToString.PadLeft(2), kv.Value.シート名, kv.Key))
                                            ret = False
                                            If cnt = max Then Return ret
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next

        Return ret
    End Function

    ''' <summary>
    ''' 再計算処理を行う
    ''' </summary>
    ''' <param name="dcKobetsu"></param>
    ''' <param name="dtCreateRonri"></param>
    ''' <param name="progressDialog"></param>
    ''' <remarks></remarks>
    Private Sub Recalculate(ByVal dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目), ByVal dtCreateRonri As DataTable, ByVal progressDialog As ProgressDialog)
        Dim dtChoItemMst As DataTable
        Dim dtKobetsuItemMst As DataTable
        Dim kobetsuList As Dictionary(Of String, Object)
        Dim itemInfoList As List(Of CreateKobetsu.ItemInfo)
        Dim ItemInfoListKobetsu As List(Of CreateKobetsu.ItemInfo)

        Try

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
                '調査票項目マスタ取得
                'REV-001 START-----------
                dtChoItemMst = DAOOther.GetChosahyoItemMaster(db, CommonInfo.Chosakubun, _pKey.chosaNen)
                'REV-001END----------
                '個別結果表項目マスタ取得(裏項番含める)
                ' REV_002↓
                'dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, True)
                dtKobetsuItemMst = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(_pKey.chosaNen, CommonInfo.Chosakubun), True)
                ' REV_002↑

                kobetsuList = New Dictionary(Of String, Object)
                For Each kv As KeyValuePair(Of String, DAOKobetsuKekkahyo.個別結果表項目) In dcKobetsu
                    kobetsuList.Add(kv.Key, kv.Value.値)
                Next

                '個別結果表・個別結果検討表作成クラス
                Dim kobetsu As CreateKobetsu = New CreateKobetsu(db,
                                                                 CommonInfo.Chosakubun,
                                                                 _pKey.chosaNen,
                                                                 CreateKobetsu.enmCreateType.個別結果表再計算,
                                                                 dtChoItemMst,
                                                                 dtKobetsuItemMst,
                                                                 Nothing,
                                                                 dtCreateRonri,
                                                                 kobetsuList,
                                                                 progressDialog)
                '個別結果表再計算実行
                itemInfoList = kobetsu.Execute(_pKey.censusNo)
                '個別結果表(当年データ、裏項番以外、再計算対象)で抽出
                ItemInfoListKobetsu = (From n In itemInfoList Where n.ItemType = CreateKobetsu.enmItemType.個別結果表 And Not n.ItemNo.Contains("前") And n.IsHidden = False And n.IsReCalc).ToList
                For Each info In ItemInfoListKobetsu
                    '個別結果表シートデータを上書き
                    If dcKobetsu.ContainsKey(info.ItemNo) Then
                        dcKobetsu(info.ItemNo).値 = If(info.Value Is Nothing, Nothing, info.Value.ToString)
                    End If
                Next

                '個別結果表シートデータ設定
                ComUtil.KobetsuKekkahyo.SetSheetData(dcKobetsu, xlSheets, CType(Me, ComObjectProcess))
            End Using

        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class
