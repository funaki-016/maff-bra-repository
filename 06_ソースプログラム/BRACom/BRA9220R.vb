Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2022.11.21 |Daiko               | 要件No.1 制度受取金・積立金等項目名称条件処理追加
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 個別結果表出力
''' </summary>
''' <remarks></remarks>
Public Class BRA9220R
    Inherits ExcelOutputMultipleBaseClass

    ''' <summary>欠測値補完</summary>
    Private _kessokuchiHokan As String

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
    ''' <param name="kessokuchiHokan"></param>
    ''' <param name="chosaNen"></param>
    'Public Sub New(ByVal outDir As String, ByVal kessokuchiHokan As String)
    Public Sub New(ByVal outDir As String, ByVal kessokuchiHokan As String, ByVal chosaNen As String)
        'MyBase.New(ComConst.個別結果表.出力用ファイル名称.リスト(CommonInfo.Chosakubun).tempFileName, True, False, ComConst.個別結果表.出力用ファイル名称.リスト(CommonInfo.Chosakubun).reportName, outDir, True)
        MyBase.New(ComConst.個別結果表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun))).tempFileName, True, False, ComConst.個別結果表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun))).reportName, outDir, True)
        ' REV_001↑
        _kessokuchiHokan = kessokuchiHokan
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(Of T As Class)(xlSheets As Excel.Sheets, unit As T)

        Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = CType(DirectCast(unit, Object), DAOKobetsuKekkahyo.PrimaryKey)
        ' REV_001↓
        'Dim fileName As String = ComConst.個別結果表.出力用ファイル名称.リスト(CommonInfo.Chosakubun).reportName _
        Dim fileName As String = ComConst.個別結果表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _  ' REV_001↑
                                 & "_" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsx"

        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        Dim dtItem As DataTable
        Dim dtKobetsu As Dictionary(Of String, DataTable)
        Dim dcKobetsu As Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)
        Dim nousanFlg As Boolean = False
        Dim seidoUketoriChosaNen As String = Nothing

        'REV_002 START---------------
        '農産物生産費かつR4体系の場合、制度受取金・積立金等項目フラグ、調査年を設定する
        If CommonInfo.Kubun2 = ComConst.区分２.農産物生産費 And ComConst.バージョン区分.結果表等項目2022 = ComUtil.getVersionKubunTaikei(pkey.chosaNen, CommonInfo.Chosakubun) Then
            seidoUketoriChosaNen = pkey.chosaNen
            nousanFlg = True
        End If
        'REV_002 END---------------

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '個別結果表項目マスタ取得
            ' REV_001↓
            'dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun)
            dtItem = DAOOther.GetKobetsuKekkahyoItemMaster(db, CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(pkey.chosaNen, CommonInfo.Chosakubun), False, seidoUketoriChosaNen) ' REV_002
            ' REV_001 REV_002↑

            '個別結果表テーブル取得
            dtKobetsu = DAOKobetsuKekkahyo.GetTable(db, pkey, _kessokuchiHokan)
        End Using

        '進捗加増
        Me.ProgressAddValue = 1

        '個別結果表項目取得
        dcKobetsu = ComUtil.KobetsuKekkahyo.GetItem(dtItem, dtKobetsu, nousanFlg) ' REV_002

        '進捗加増
        Me.ProgressAddValue = 1

        '調査票シートデータ設定
        ComUtil.KobetsuKekkahyo.SetSheetData(dcKobetsu, xlSheets, CType(Me, ComObjectProcess))

        '進捗加増
        Me.ProgressAddValue = 1
    End Sub
End Class
