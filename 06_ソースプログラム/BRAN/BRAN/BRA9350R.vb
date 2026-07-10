Imports Microsoft.Office.Interop

''' <summary>
''' 集計結果検討表（報告用）出力
''' </summary>
''' <remarks></remarks>
Public Class BRA9350R
    Inherits ExcelOutputMultipleBaseClass

    ''' <summary>調査区分</summary>
    Private _chosakubun As String

    ''' <summary>集計結果検討表作成オブジェクト</summary>
    Private _kentohyo As CreateSyukeiKentohyo

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
    ''' <param name="chosakubun"></param>
    ''' <param name="kentohyo"></param>
    ''' <param name="chosaNen"></param>
    Public Sub New(ByVal outDir As String, chosakubun As String, kentohyo As CreateSyukeiKentohyo, ByVal chosaNen As String)
        MyBase.New(ComConst.集計結果検討表_報告用.出力用ファイル名称.リスト(Tuple.Create(chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, chosakubun))).tempFileName, True, False, ComConst.集計結果検討表_報告用.出力用ファイル名称.リスト(Tuple.Create(chosakubun, ComUtil.getVersionKubunTaikei(chosaNen, chosakubun))).reportName, outDir, True)
        _chosakubun = chosakubun
        _kentohyo = kentohyo
    End Sub

    ''' <summary>
    ''' 帳票編集
    ''' </summary>
    ''' <param name="xlSheets"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(Of T As Class)(xlSheets As Excel.Sheets, unit As T)

        Dim keyPair As BRA9330F.KeyPair = CType(DirectCast(unit, Object), BRA9330F.KeyPair)
        Dim pkeyThis As DAOSyukeiKekkahyo.PrimaryKey = keyPair.PKeyThis
        Dim pkeyLast As DAOSyukeiKekkahyo.PrimaryKey = keyPair.PKeyLast
        Dim fileName As String = ComConst.集計結果検討表_報告用.出力用ファイル名称.リスト(Tuple.Create(_chosakubun, ComUtil.getVersionKubunTaikei(pkeyThis.chosaNen, _chosakubun))).reportName _
                                 & "_" & pkeyThis.chosaNen & "_" & pkeyThis.syukeiNo & ".xlsx"

        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        Dim dtSyukei As Dictionary(Of String, DataTable)
        Dim dcSyukei As Dictionary(Of String, ComUtil.SyukeiKekkaKentohyo.集計結果検討表項目)

        Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))
            '集計結果表テーブル取得
            dtSyukei = DAOSyukeiKekkahyo.GetTable(db, _chosakubun, pkeyThis)
        End Using

        Dim kkey As DAOSyukeiKekkahyo.KomokuKey = Nothing
        If dtSyukei(ComConst.集計結果表.テーブル名称(_chosakubun)(0)).Rows.Count > 0 Then
            Dim dr As DataRow = dtSyukei(ComConst.集計結果表.テーブル名称(_chosakubun)(0))(0)
            fileName = ComConst.集計結果検討表_報告用.出力用ファイル名称.リスト(Tuple.Create(_chosakubun, ComUtil.getVersionKubunTaikei(pkeyThis.chosaNen, _chosakubun))).reportName _
                       & "_" & dr("調査年").ToString _
                       & "_" & dr("集計番号").ToString _
                       & "_" & dr("生産費平均値種類").ToString _
                       & "_" & Integer.Parse(dr("地域コード").ToString).ToString("00") _
                       & "_" & Integer.Parse(dr("規模階層").ToString).ToString("00") _
                       & "_" & dr("田畑区分").ToString _
                       & "_" & dr("ビール麦販売区分").ToString _
                       & "_" & dr("てんさい栽培区分").ToString & ".xlsx"

            kkey = New DAOSyukeiKekkahyo.KomokuKey(CommonInfo.Kyoku, CommonInfo.Jimusyo, CommonInfo.Center, dr("平均種類").ToString, dr("規模階層").ToString, _
                                                   dr("生産費平均値種類").ToString, dr("田畑区分").ToString, dr("ビール麦販売区分").ToString, dr("てんさい栽培区分").ToString, _
                                                   dr("集計名称").ToString, dr("規模階層").ToString, dr("地域コード").ToString, dr("部門コード").ToString)

        End If

        Me.OutPath = IO.Path.Combine(Me.OutDir, fileName)

        '進捗加増
        Me.ProgressAddValue = 1

        Dim itemInfoList As List(Of CreateSyukeiKentohyo.ItemInfo)
        Dim ItemInfoListKento As List(Of CreateSyukeiKentohyo.ItemInfo)
        '集計結果検討表項目取得
        If ComUtil.IsChikusan() OrElse ComUtil.IsNousan() Then
            itemInfoList = _kentohyo.Execute(pkeyThis, pkeyLast, kkey)
        Else
            itemInfoList = _kentohyo.Execute(pkeyThis, pkeyLast)
        End If
        '集計結果検討表(当年データ、裏項番以外)で抽出
        ItemInfoListKento = (From n In itemInfoList Where n.ItemType = CreateSyukeiKentohyo.enmItemType.集計結果検討表 And Not n.ItemNo.Contains("前") And n.IsHidden = False).ToList
        dcSyukei = New Dictionary(Of String, ComUtil.SyukeiKekkaKentohyo.集計結果検討表項目)
        For Each info In ItemInfoListKento
            Dim item As New ComUtil.SyukeiKekkaKentohyo.集計結果検討表項目
            With item
                .シート名 = info.SheetName
                .行位置 = info.Row
                .列位置 = info.Col
                .値 = If(info.Value Is Nothing, Nothing, info.Value.ToString)
                .型区分 = info.ValueType
            End With
            dcSyukei.Add(info.ItemNo, item)
        Next

        '進捗加増
        Me.ProgressAddValue = 1

        '集計結果検討表シートデータ設定
        ComUtil.SyukeiKekkaKentohyo.SetSheetData(dcSyukei, xlSheets, CType(Me, ComObjectProcess))

        '進捗加増
        Me.ProgressAddValue = 1
    End Sub
End Class
