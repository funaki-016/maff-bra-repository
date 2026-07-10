Imports Microsoft.Office.Interop

'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_001   | 2022.10.11 |Daiko               | 要件No1 バージョン区分追加
'//  REV_002   | 2025.09.10 |GCU                 | 要件No3 個別結果検討表ファイル名への作成日時の挿入
'//            |            |                    |
'//*************************************************************************************************
''' <summary>
''' 個別結果検討表出力
''' </summary>
''' <remarks></remarks>
Public Class BRA9230R
    Inherits ExcelOutputMultipleBaseClass

    ''' <summary>個別結果表・個別結果検討表作成オブジェクト</summary>
    Private _kobetsu As CreateKobetsu

    ' REV_001↓
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="outDir"></param>
    ''' <param name="kobetsu"></param>
    ''' <param name="chosaNen"></param>
    'Public Sub New(ByVal outDir As String, ByVal kobetsu As CreateKobetsu)
    Public Sub New(ByVal outDir As String, ByVal kobetsu As CreateKobetsu, ByVal chosaNen As String)
        'MyBase.New(ComConst.個別結果検討表.出力用ファイル名称.リスト(CommonInfo.Chosakubun).tempFileName, True, False, ComConst.個別結果検討表.出力用ファイル名称.リスト(CommonInfo.Chosakubun).reportName, outDir, True)
        MyBase.New(ComConst.個別結果検討表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun))).tempFileName, True, False, ComConst.個別結果検討表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, CommonInfo.Chosakubun))).reportName, outDir, True)
        ' REV_001↑
        _kobetsu = kobetsu
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(Of T As Class)(xlSheets As Excel.Sheets, unit As T)

        Dim pkey As DAOKobetsuKekkahyo.PrimaryKey = CType(DirectCast(unit, Object), DAOKobetsuKekkahyo.PrimaryKey)
        ' REV_001↓REV_002↓
        'Dim fileName As String = ComConst.個別結果検討表.出力用ファイル名称.リスト(CommonInfo.Chosakubun).reportName _
        Dim fileName As String
        If CommonInfo.Kubun2 = ComConst.区分２.営農類型別経営統計 Then
            fileName = ComConst.個別結果検討表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(pkey.chosaNen, CommonInfo.Chosakubun))).reportName _
                         & "_" & pkey.chosaNen & "_" & pkey.censusNo & ".xlsx"

        Else
            fileName = ComConst.個別結果検討表.出力用ファイル名称.リスト(Tuple.Create(CommonInfo.Chosakubun, ComUtil.getVersionKubunTaikei(
                        pkey.chosaNen, CommonInfo.Chosakubun))).reportName & "_" & pkey.chosaNen & "_" & pkey.censusNo & "_" & DateTime.Now.ToString("yyyyMMddHHmm") & ".xlsx"
        End If

        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        Dim dcKobetsu As Dictionary(Of String, ComUtil.KobetsuKekkaKentohyo.個別結果検討表項目)

        '進捗加増
        Me.ProgressAddValue = 1

        Dim itemInfoList As List(Of CreateKobetsu.ItemInfo)
        Dim ItemInfoListKento As List(Of CreateKobetsu.ItemInfo)
        '個別結果検討表項目取得
        itemInfoList = _kobetsu.Execute(pkey.censusNo)
        '個別結果検討表(当年データ、裏項番以外)で抽出
        ItemInfoListKento = (From n In itemInfoList Where n.ItemType = CreateKobetsu.enmItemType.個別結果検討表 And Not n.ItemNo.Contains("前") And n.IsHidden = False).ToList
        dcKobetsu = New Dictionary(Of String, ComUtil.KobetsuKekkaKentohyo.個別結果検討表項目)
        For Each info In ItemInfoListKento
            Dim item As New ComUtil.KobetsuKekkaKentohyo.個別結果検討表項目
            With item
                .シート名 = info.SheetName
                .行位置 = info.Row
                .列位置 = info.Col
                .値 = If(info.Value Is Nothing, Nothing, info.Value.ToString)
                .型区分 = info.ValueType
            End With
            dcKobetsu.Add(info.ItemNo, item)
        Next

        '進捗加増
        Me.ProgressAddValue = 1

        '個別結果検討表シートデータ設定
        ComUtil.KobetsuKekkaKentohyo.SetSheetData(dcKobetsu, xlSheets, CType(Me, ComObjectProcess))

        '進捗加増
        Me.ProgressAddValue = 1
    End Sub
End Class
