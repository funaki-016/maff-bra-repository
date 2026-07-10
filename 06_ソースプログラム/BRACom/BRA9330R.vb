Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.12.15 |Daiko               | 要件No4 バージョン区分追加
'//  REV_002   | 2023.01.11 |Daiko               | 要件No.4 制度受取金・積立金等項目名称取得処理追加
'//  REV_003   | 2023.03.14 |Daiko               | 要件No.12 牛乳の生産費平均値種類が0以外の場合はNo3シートを非表示化
'//*************************************************************************************************
''' <summary>
''' 集計結果表出力
''' </summary>
''' <remarks></remarks>
Public Class BRA9330R
    Inherits ExcelOutputMultipleBaseClass

    ''' <summary>調査区分</summary>
    Private _chosakubun As String

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ' REV_001↓
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <param name="chosakubun"></param>
    ''' <param name="chosaNen"></param>
    'Public Sub New(ByVal outDir As String, chosakubun As String)
    Public Sub New(ByVal outDir As String, chosakubun As String, ByVal chosaNen As String)
        'MyBase.New(ComConst.集計結果表.出力用ファイル名称.リスト(chosakubun).tempFileName, True, False, ComConst.集計結果表.出力用ファイル名称.リスト(chosakubun).reportName, outDir, True)
        MyBase.New(ComConst.集計結果表.出力用ファイル名称.リスト(Tuple.Create(chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, chosakubun))).tempFileName, True, False, ComConst.集計結果表.出力用ファイル名称.リスト(Tuple.Create(chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, chosakubun))).reportName, outDir, True)
        ' REV_001↑
        _chosakubun = chosakubun
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(Of T As Class)(xlSheets As Excel.Sheets, unit As T)

        Dim pkey As DAOSyukeiKekkahyo.PrimaryKey = CType(DirectCast(unit, Object), DAOSyukeiKekkahyo.PrimaryKey)
        ' REV_001↓
        'Dim fileName As String = ComConst.集計結果表.出力用ファイル名称.リスト(_chosakubun).reportName _
        '                         & "_" & pkey.chosaNen & "_" & pkey.syukeiNo & ".xlsx"
        Dim fileName As String = ComConst.集計結果表.出力用ファイル名称.リスト(Tuple.Create(_chosakubun, ComUtil.getVersionKubunTaikei(pkey.chosaNen, _chosakubun))).reportName _
                                 & "_" & pkey.chosaNen & "_" & pkey.syukeiNo & ".xlsx"
        ' REV_001↑

        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        Dim dtItem As DataTable = Nothing
        Dim dtSyukei As Dictionary(Of String, DataTable)
        Dim dcSyukei As Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)
        Dim seidoUketoriChosaNen As String = Nothing 'REV_002
        Dim nousanFlg As Boolean = False 'REV_002
        Dim hiddenSheets = New List(Of String) 'REV_003

        'REV_002 START---------------
        '農産物生産費かつR4体系の場合
        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 And ComConst.バージョン区分.結果表等項目2022 = ComUtil.getVersionKubunTaikei(pkey.chosaNen, CommonInfo.Chosakubun) Then
            seidoUketoriChosaNen = pkey.chosaNen
            nousanFlg = True
        End If
        'REV_002 END---------------

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '集計結果表テーブル取得
            dtSyukei = DAOSyukeiKekkahyo.GetTable(db, _chosakubun, pkey)

            If dtSyukei(ComConst.集計結果表.テーブル名称(_chosakubun)(0)).Rows.Count > 0 Then
                Dim dr As DataRow = dtSyukei(ComConst.集計結果表.テーブル名称(_chosakubun)(0))(0)
                ' REV_001↓
                'fileName = ComConst.集計結果表.出力用ファイル名称.リスト(_chosakubun).reportName _
                '           & "_" & dr("調査年").ToString _
                '           & "_" & dr("集計番号").ToString _
                '           & "_" & dr("生産費平均値種類").ToString _
                '           & "_" & Integer.Parse(dr("地域コード").ToString).ToString("00") _
                '           & "_" & Integer.Parse(dr("規模階層").ToString).ToString("00") _
                '           & "_" & dr("田畑区分").ToString _
                '           & "_" & dr("ビール麦販売区分").ToString _
                '           & "_" & dr("てんさい栽培区分").ToString & ".xlsx"
                fileName = ComConst.集計結果表.出力用ファイル名称.リスト(Tuple.Create(_chosakubun, ComUtil.getVersionKubunTaikei(pkey.chosaNen, _chosakubun))).reportName _
                           & "_" & dr("調査年").ToString _
                           & "_" & dr("集計番号").ToString _
                           & "_" & dr("生産費平均値種類").ToString _
                           & "_" & Integer.Parse(dr("地域コード").ToString).ToString("00") _
                           & "_" & Integer.Parse(dr("規模階層").ToString).ToString("00") _
                           & "_" & dr("田畑区分").ToString _
                           & "_" & dr("ビール麦販売区分").ToString _
                           & "_" & dr("てんさい栽培区分").ToString & ".xlsx"
                ' REV_001↑

                '集計結果表項目マスタ取得
                ' REV_001↓
                'dtItem = DAOOther.GetSyukeiKekkahyoItemMaster(db, _chosakubun, dr("生産費平均値種類").ToString)
                dtItem = DAOOther.GetSyukeiKekkahyoItemMaster(db, _chosakubun, ComUtil.getVersionKubunTaikei(pkey.chosaNen, _chosakubun), dr("生産費平均値種類").ToString, False, False, seidoUketoriChosaNen) ' REV_002
                ' REV_001,REV_002↑
                ' REV_003 ↓
                '牛乳の生産費平均値種類が0以外の場合はNo3シートを非表示化
                If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 _
                    AndAlso dr("生産費平均値種類").ToString <> "0" Then
                    hiddenSheets.Add("No3")
                End If
                ' REV_003 ↑
            End If
        End Using

        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        '進捗加増
        Me.ProgressAddValue = 1

        '集計結果表項目取得
        dcSyukei = ComUtil.SyukeiKekkahyo.GetItem(dtItem, dtSyukei, nousanFlg) ' REV_002

        '進捗加増
        Me.ProgressAddValue = 1
        '集計結果表シートデータ設定
        ' REV_003↓
        'ComUtil.SyukeiKekkahyo.SetSheetData(dcSyukei, xlSheets, CType(Me, ComObjectProcess))
        ComUtil.SyukeiKekkahyo.SetSheetData(dcSyukei, xlSheets, CType(Me, ComObjectProcess), hiddenSheets)
        ' REV_003↑

        '進捗加増
        Me.ProgressAddValue = 1
    End Sub
End Class
