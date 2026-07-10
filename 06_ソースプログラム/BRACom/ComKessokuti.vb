Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.Office.Interop

''' <summary>
''' 欠測値テーブル操作
''' </summary>
''' <remarks></remarks>
Partial Public Class ComUtil

#Region "画面描画処理"

    ''' <summary>
    ''' 調査年コンボボックス設定
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="tableName"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetChosaNenComboBox(ByRef chosaNenComboBox As ComboBox, ByVal tableName As String)

        chosaNenComboBox.ValueMember = "調査年"
        chosaNenComboBox.DisplayMember = "調査年"

        Using db As New DBAccess(My.Settings.BRAHConnectionString)
            chosaNenComboBox.DataSource = DAOOther.GetChosaNen(db, tableName)
        End Using

    End Sub

    ''' <summary>
    ''' 更新日時設定
    ''' </summary>
    ''' <param name="updateTextBox"></param>
    ''' <param name="chosaNen"></param>
    ''' <param name="tableName"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateDate(ByRef updateTextBox As TextBox, ByVal chosaNen As String, ByVal tableName As String)

        Dim dtn = String.Empty

        Using db As New DBAccess(My.Settings.BRAHConnectionString)
            dtn = DAOOther.GetUpdateDate(db, chosaNen, tableName)
        End Using

        updateTextBox.Text = If(Not String.IsNullOrWhiteSpace(dtn), DateTime.Parse(dtn).ToString(ComConst.DATETIME_FORMAT), String.Empty)

    End Sub

#End Region

#Region "データベース操作"

    ''' <summary>
    ''' 営農欠測値適用平均値データテーブルクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 営農欠測値適用平均値データ
        Public 調査年 As Integer
        Public 農業地域 As Integer
        Public 営農類型 As Integer
        Public 営農規模 As Integer
        Public 集計数 As Decimal?
        Public 現金 As Decimal?
        Public 預金等 As Decimal?
        Public 売掛金＿未収金 As Decimal?
        Public 自動車＿農機具 As Decimal?
        Public 建物＿構築物 As Decimal?
        Public 土地 As Decimal?
        Public 果樹＿牛馬等 As Decimal?
        Public 流動負債 As Decimal?
        Public 買掛金 As Decimal?
        Public 短期借入金 As Decimal?
        Public 長期借入金 As Decimal?
        Public 更新日付 As Date?
        Public 更新者ID As String
    End Class

    ''' <summary>
    ''' 営農欠測値平均値代入結果テーブルクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class 営農欠測値平均値代入結果
        Public 調査年 As Integer
        Public センサス番号 As String
        Public 営農類型 As Decimal?
        Public 営農規模 As Decimal?
        Public 適用＿営農類型 As Decimal?
        Public 適用＿営農規模 As Decimal?
        Public 適用＿農業地域 As Decimal?
        Public 適用＿不可 As Decimal?
        Public 更新日付 As Date?
        Public 更新者ID As String
    End Class

    ''' <summary>
    ''' 「営農欠測値適用平均値データ」テーブルからレコードを抽出する
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="whereAnd"></param>
    ''' <param name="orderBy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function GetAverageDataRowList(ByVal chosaNen As String,
                                          Optional ByVal whereAnd As String = "",
                                          Optional ByVal orderBy As String = "") As List(Of 営農欠測値適用平均値データ)

        If String.IsNullOrWhiteSpace(chosaNen) Then
            Return Nothing
        End If

        Dim dataTable As DataTable = Nothing

        Using db As New DBAccess(My.Settings.BRAHConnectionString)
            dataTable = DAOOther.SelectAverageDataTable(db, chosaNen, whereAnd, orderBy)
        End Using

        Dim rowList As New List(Of 営農欠測値適用平均値データ)

        For Each dataRow As DataRow In dataTable.Rows

            Dim row As New 営農欠測値適用平均値データ

            CastRow(dataRow, row)

            rowList.Add(row)
        Next

        Return rowList

    End Function

    ''' <summary>
    ''' 「営農欠測値平均値代入結果」テーブルからレコードを抽出する
    ''' </summary>
    ''' <param name="chosaNen"></param>
    ''' <param name="whereAnd"></param>
    ''' <param name="orderBy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function GetAverageResultTableRowList(ByVal chosaNen As String,
                                                 Optional ByVal whereAnd As String = "",
                                                 Optional ByVal orderBy As String = "") As List(Of 営農欠測値平均値代入結果)

        If String.IsNullOrWhiteSpace(chosaNen) Then
            Return Nothing
        End If

        Dim dataTable As DataTable = Nothing

        Using db As New DBAccess(My.Settings.BRAHConnectionString)
            dataTable = DAOOther.SelectAverageResultTable(db, chosaNen, whereAnd, orderBy)
        End Using

        Dim rowList As New List(Of 営農欠測値平均値代入結果)

        For Each dataRow As DataRow In dataTable.Rows

            Dim row As New 営農欠測値平均値代入結果

            CastRow(dataRow, row)

            rowList.Add(row)
        Next

        Return rowList

    End Function

    ''' <summary>
    ''' DBから取得したデータをエンティティクラスに変換する
    ''' </summary>
    ''' <param name="dataRow"></param>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Shared Sub CastRow(ByVal dataRow As DataRow, ByRef row As 営農欠測値適用平均値データ)

        row.調査年 = CInt(dataRow.Item("調査年"))
        row.農業地域 = CInt(dataRow.Item("農業地域"))
        row.営農類型 = CInt(dataRow.Item("営農類型"))
        row.営農規模 = CInt(dataRow.Item("営農規模"))
        row.集計数 = GetNullOrDec(dataRow.Item("集計数"))
        row.現金 = GetNullOrDec(dataRow.Item("現金"))
        row.預金等 = GetNullOrDec(dataRow.Item("預金等"))
        row.売掛金＿未収金 = GetNullOrDec(dataRow.Item("売掛金＿未収金"))
        row.自動車＿農機具 = GetNullOrDec(dataRow.Item("自動車＿農機具"))
        row.建物＿構築物 = GetNullOrDec(dataRow.Item("建物＿構築物"))
        row.土地 = GetNullOrDec(dataRow.Item("土地"))
        row.果樹＿牛馬等 = GetNullOrDec(dataRow.Item("果樹＿牛馬等"))
        row.流動負債 = GetNullOrDec(dataRow.Item("流動負債"))
        row.買掛金 = GetNullOrDec(dataRow.Item("買掛金"))
        row.短期借入金 = GetNullOrDec(dataRow.Item("短期借入金"))
        row.長期借入金 = GetNullOrDec(dataRow.Item("長期借入金"))
        row.更新日付 = GetNullOrDateConvertString(dataRow.Item("更新日付"))
        row.更新者ID = GetBlankOrString(dataRow.Item("更新者ID"))

    End Sub

    ''' <summary>
    ''' DBから取得したデータをエンティティクラスに変換する
    ''' </summary>
    ''' <param name="dataRow"></param>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Shared Sub CastRow(ByVal dataRow As DataRow, ByRef row As 営農欠測値平均値代入結果)

        row.調査年 = CInt(dataRow.Item("調査年"))
        row.センサス番号 = GetBlankOrString(dataRow.Item("センサス番号"))
        row.営農類型 = GetNullOrDec(dataRow.Item("営農類型"))
        row.営農規模 = GetNullOrDec(dataRow.Item("営農規模"))
        row.適用＿営農類型 = GetNullOrDec(dataRow.Item("適用＿営農類型"))
        row.適用＿営農規模 = GetNullOrDec(dataRow.Item("適用＿営農規模"))
        row.適用＿農業地域 = GetNullOrDec(dataRow.Item("適用＿農業地域"))
        row.適用＿不可 = GetNullOrDec(dataRow.Item("適用＿不可"))
        row.更新日付 = GetNullOrDateConvertString(dataRow.Item("更新日付"))
        row.更新者ID = GetBlankOrString(dataRow.Item("更新者ID"))

    End Sub

#End Region

#Region "エクセル操作"

    ''' <summary>
    ''' シートデータ設定
    ''' </summary>
    ''' <param name="rowList"></param>
    ''' <param name="xlSheets"></param>
    ''' <param name="xlsProcess"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSheetData(ByVal rowList As List(Of 営農欠測値適用平均値データ), xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)

        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.欠測値.適用データ一覧表出力用ファイル.Report.SheetName), Excel.Worksheet)

            'シート保護確認
            Dim protect As Boolean = xlSheet.ProtectContents
            If protect Then
                xlSheet.Unprotect()
            End If

            Dim rng As Excel.Range = Nothing
            Try
                '明細一覧
                Dim arrData(,) As Object

                rng = xlSheet.Range(ComConst.欠測値.適用データ一覧表出力用ファイル.Col.First & ComConst.欠測値.適用データ一覧表出力用ファイル.Row.First & ":" _
                                    & ComConst.欠測値.適用データ一覧表出力用ファイル.Col.Last & rowList.Count + ComConst.欠測値.適用データ一覧表出力用ファイル.Row.First - 1)

                arrData = DirectCast(rng.Formula, Object(,))

                Dim i As Integer = 1

                'DBから取得したレコード数ぶんループする
                rowList.ForEach(
                    Sub(row)
                        arrData(i, 1) = row.農業地域
                        arrData(i, 2) = row.営農類型
                        arrData(i, 3) = row.営農規模
                        arrData(i, 4) = row.集計数
                        arrData(i, 5) = row.現金
                        arrData(i, 6) = row.預金等
                        arrData(i, 7) = row.売掛金＿未収金
                        arrData(i, 8) = row.自動車＿農機具
                        arrData(i, 9) = row.建物＿構築物
                        arrData(i, 10) = row.土地
                        arrData(i, 11) = row.果樹＿牛馬等
                        arrData(i, 12) = row.流動負債
                        arrData(i, 13) = row.買掛金
                        arrData(i, 14) = row.短期借入金
                        arrData(i, 15) = row.長期借入金

                        i = i + 1
                    End Sub)

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
    ''' シートデータ設定
    ''' </summary>
    ''' <param name="rowList"></param>
    ''' <param name="xlSheets"></param>
    ''' <param name="xlsProcess"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetSheetData(ByVal rowList As List(Of 営農欠測値平均値代入結果), xlSheets As Excel.Sheets, xlsProcess As ExcelProcess)

        Dim xlSheet As Excel.Worksheet = Nothing

        Try
            'シートの設定
            xlSheet = DirectCast(xlSheets.Item(ComConst.欠測値.適用状況一覧表出力用ファイル.Report.SheetName), Excel.Worksheet)

            'シート保護確認
            Dim protect As Boolean = xlSheet.ProtectContents
            If protect Then
                xlSheet.Unprotect()
            End If

            Dim rng As Excel.Range = Nothing
            Try
                '明細一覧
                Dim arrData(,) As Object

                rng = xlSheet.Range(ComConst.欠測値.適用状況一覧表出力用ファイル.Col.First & ComConst.欠測値.適用状況一覧表出力用ファイル.Row.First & ":" _
                                    & ComConst.欠測値.適用状況一覧表出力用ファイル.Col.Last & rowList.Count + ComConst.欠測値.適用状況一覧表出力用ファイル.Row.First - 1)

                arrData = DirectCast(rng.Formula, Object(,))

                Dim i As Integer = 1

                'DBから取得したレコード数ぶんループする
                rowList.ForEach(
                    Sub(row)
                        arrData(i, 1) = row.センサス番号
                        arrData(i, 2) = ""
                        arrData(i, 3) = row.営農類型
                        arrData(i, 4) = row.営農規模
                        arrData(i, 5) = row.適用＿営農類型
                        arrData(i, 6) = row.適用＿営農規模
                        arrData(i, 7) = row.適用＿農業地域
                        arrData(i, 8) = row.適用＿不可

                        i = i + 1
                    End Sub)

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

#End Region

End Class
