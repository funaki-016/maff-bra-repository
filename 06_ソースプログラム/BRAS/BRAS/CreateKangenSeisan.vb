Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Globalization

Public Class CreateKangenSeisan
    Inherits BRA7420R

    Public Class 生産費

        Public Class sheet1
            Public Const WAREKI_CELL As String = "O11"
            Public Const KUBUN_CELL As String = "J16"
            Public Const PICTURE_CELL As String = "G19"
            Public Const AISATSU_CELL As String = "E46"
            Public Const SHUKEI_CELL As String = "BA11:BC14"

        End Class

        Public Class sheet2
            Public Class CellAddress
                Public PER1 As String
                Public PER2 As String
                Public PER3 As String

                Public THIS_YEAR_1 As String
                Public PREV_YEAR_1 As String
                Public THIS_YEAR_2 As String
                Public PREV_YEAR_2 As String
                Public THIS_YEAR_3 As String
                Public PREV_YEAR_3 As String

                Public KOBETSU_DATA1 As String
                Public KOBETSU_DATA2 As String
                Public KOBETSU_DATA3 As String
            End Class

            Public Class Define
                Public DATA1_ROW As Integer
                Public DATA2_ROW As Integer
                Public DATA3_ROW As Integer
            End Class

            Public Class SumInfo
                Public ROW As Integer
                Public KOBETSU_KOBAN1 As String
                Public KOBETSU_KOBAN2 As String
            End Class

            Public Class sumInfo2
                Public ROW As Integer
                Public COL As Integer
                Public KOBETSU_KOBAN1 As String
                Public KOBETSU_KOBAN2 As String
                Public KOBETSU_KOBAN3 As String
                Public KOBETSU_KOBAN4 As String
            End Class
            Public Shared CellList As New Dictionary(Of String, CellAddress) From {
                  {ComConst.調査区分.米生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I28", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I29", .PREV_YEAR_3 = "O29", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M24", .KOBETSU_DATA3 = "I31:R34"}} _
                , {ComConst.調査区分.小麦生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.二条大麦生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.六条大麦生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.はだか麦生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.そば生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.大豆生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.原料用かんしょ生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.なたね生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.てんさい生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.さとうきび生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.米生産費統計_組織法人, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I28", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I29", .PREV_YEAR_3 = "O29", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M24", .KOBETSU_DATA3 = "I31:R34"}} _
                , {ComConst.調査区分.小麦生産費統計_組織法人, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I29", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I30", .PREV_YEAR_3 = "O30", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M25", .KOBETSU_DATA3 = "I32:R35"}} _
                , {ComConst.調査区分.大豆生産費統計_組織法人, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I29", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I30", .PREV_YEAR_3 = "O30", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M25", .KOBETSU_DATA3 = "I32:R35"}} _
                , {ComConst.調査区分.牛乳生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I20", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I21", .PREV_YEAR_3 = "O21", .KOBETSU_DATA1 = "D9:E39", .KOBETSU_DATA2 = "I9:M16", .KOBETSU_DATA3 = "I23:R28"}} _
                , {ComConst.調査区分.子牛生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E39", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R24"}} _
                , {ComConst.調査区分.乳用雄育成牛生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R25"}} _
                , {ComConst.調査区分.交雑種育成牛生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R25"}} _
                , {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R25"}} _
                , {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R25"}} _
                , {ComConst.調査区分.交雑種肥育牛生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R25"}} _
                , {ComConst.調査区分.肥育豚生産費統計_個別, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E40", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                , {ComConst.調査区分.経営分析調査_二条大麦生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_六条大麦生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_はだか麦生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_そば生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_なたね生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_てんさい生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_さとうきび生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I31", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I32", .PREV_YEAR_3 = "O32", .KOBETSU_DATA1 = "D9:E37", .KOBETSU_DATA2 = "I9:M27", .KOBETSU_DATA3 = "I34:R37"}} _
                , {ComConst.調査区分.経営分析調査_牛乳生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I19", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I20", .PREV_YEAR_3 = "O20", .KOBETSU_DATA1 = "D9:E39", .KOBETSU_DATA2 = "I9:M15", .KOBETSU_DATA3 = "I22:R27"}} _
                , {ComConst.調査区分.経営分析調査_子牛生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E39", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R23"}} _
                , {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                , {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                , {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                , {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                , {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E38", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                , {ComConst.調査区分.経営分析調査_肥育豚生産費, New CellAddress With {.PER1 = "D6", .PER2 = "I6", .PER3 = "I18", .THIS_YEAR_1 = "D8", .PREV_YEAR_1 = "E8", .THIS_YEAR_2 = "I8", .PREV_YEAR_2 = "M8", .THIS_YEAR_3 = "I19", .PREV_YEAR_3 = "O19", .KOBETSU_DATA1 = "D9:E40", .KOBETSU_DATA2 = "I9:M14", .KOBETSU_DATA3 = "I21:R24"}} _
                 }

            Public Shared DefineList As New Dictionary(Of String, Define) From {
                  {ComConst.調査区分.米生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 16, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.小麦生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.二条大麦生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.六条大麦生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.はだか麦生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.そば生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.大豆生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.原料用かんしょ生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.なたね生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.てんさい生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.さとうきび生産費統計_個別, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.米生産費統計_組織法人, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 16, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.小麦生産費統計_組織法人, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 17, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.大豆生産費統計_組織法人, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 17, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.牛乳生産費統計_個別, New Define With {.DATA1_ROW = 31, .DATA2_ROW = 8, .DATA3_ROW = 6}} _
                , {ComConst.調査区分.子牛生産費統計_個別, New Define With {.DATA1_ROW = 31, .DATA2_ROW = 7, .DATA3_ROW = 3}} _
                , {ComConst.調査区分.乳用雄育成牛生産費統計_個別, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 7, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.交雑種育成牛生産費統計_個別, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 7, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 7, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 7, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.交雑種肥育牛生産費統計_個別, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 7, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.肥育豚生産費統計_個別, New Define With {.DATA1_ROW = 32, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_二条大麦生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_六条大麦生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_はだか麦生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_そば生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_なたね生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_てんさい生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_さとうきび生産費, New Define With {.DATA1_ROW = 29, .DATA2_ROW = 19, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_牛乳生産費, New Define With {.DATA1_ROW = 31, .DATA2_ROW = 7, .DATA3_ROW = 6}} _
                , {ComConst.調査区分.経営分析調査_子牛生産費, New Define With {.DATA1_ROW = 31, .DATA2_ROW = 6, .DATA3_ROW = 3}} _
                , {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, New Define With {.DATA1_ROW = 30, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
                , {ComConst.調査区分.経営分析調査_肥育豚生産費, New Define With {.DATA1_ROW = 32, .DATA2_ROW = 6, .DATA3_ROW = 4}} _
            }

            Public Const KAISETSU1 As String = "角丸四角形吹き出し 5"
            Public Const KAISETSU2 As String = "角丸四角形吹き出し 7"

        End Class

        Public Class sheet3
            Public Class CellAddress
                Public PER1 As String
                Public PER1_SHU As String
                Public PER2 As String
                Public PER2_SHU As String
                Public PER3 As String
                Public PER3_SHU As String

                Public THIS_YEAR_1 As String
                Public PREV_YEAR_1 As String
                Public THIS_YEAR_1_SHU As String
                Public PREV_YEAR_1_SHU As String
                Public THIS_YEAR_2 As String
                Public PREV_YEAR_2 As String
                Public THIS_YEAR_2_SHU As String
                Public PREV_YEAR_2_SHU As String
                Public THIS_YEAR_3 As String
                Public PREV_YEAR_3 As String
                Public THIS_YEAR_3_SHU As String
                Public PREV_YEAR_3_SHU As String

                Public KOBETSU_DATA1 As String
                Public SHUKEI_DATA1 As String
                Public KOBETSU_DATA2 As String
                Public SHUKEI_DATA2 As String
                Public KOBETSU_DATA3 As String
                Public SHUKEI_DATA3 As String
            End Class

            Public Shared リスト As New Dictionary(Of String, CellAddress) From {
                  {ComConst.調査区分.米生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L29", .PER3_SHU = "X29", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L31", .PREV_YEAR_3 = "R31", .THIS_YEAR_3_SHU = "X31", .PREV_YEAR_3_SHU = "AD31", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P25", .SHUKEI_DATA2 = "X10:AB25", .KOBETSU_DATA3 = "L33:U36", .SHUKEI_DATA3 = "X33:AG36"}} _
                , {ComConst.調査区分.小麦生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.二条大麦生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.六条大麦生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.はだか麦生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.そば生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.大豆生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.原料用かんしょ生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.なたね生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.てんさい生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.さとうきび生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.米生産費統計_組織法人, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L29", .PER3_SHU = "X29", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L31", .PREV_YEAR_3 = "R31", .THIS_YEAR_3_SHU = "X31", .PREV_YEAR_3_SHU = "AD31", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P25", .SHUKEI_DATA2 = "X10:AB25", .KOBETSU_DATA3 = "L33:U36", .SHUKEI_DATA3 = "X33:AG36"}} _
                , {ComConst.調査区分.小麦生産費統計_組織法人, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L30", .PER3_SHU = "X30", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L32", .PREV_YEAR_3 = "R32", .THIS_YEAR_3_SHU = "X32", .PREV_YEAR_3_SHU = "AD32", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P26", .SHUKEI_DATA2 = "X10:AB26", .KOBETSU_DATA3 = "L34:U37", .SHUKEI_DATA3 = "X34:AG37"}} _
                , {ComConst.調査区分.大豆生産費統計_組織法人, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L30", .PER3_SHU = "X30", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L32", .PREV_YEAR_3 = "R32", .THIS_YEAR_3_SHU = "X32", .PREV_YEAR_3_SHU = "AD32", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P26", .SHUKEI_DATA2 = "X10:AB26", .KOBETSU_DATA3 = "L34:U37", .SHUKEI_DATA3 = "X34:AG37"}} _
                , {ComConst.調査区分.牛乳生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L21", .PER3_SHU = "X21", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L23", .PREV_YEAR_3 = "R23", .THIS_YEAR_3_SHU = "X23", .PREV_YEAR_3_SHU = "AD23", .KOBETSU_DATA1 = "D10:E40", .SHUKEI_DATA1 = "G10:H40", .KOBETSU_DATA2 = "L10:P17", .SHUKEI_DATA2 = "X10:AB17", .KOBETSU_DATA3 = "L25:U30", .SHUKEI_DATA3 = "X25:AG30"}} _
                , {ComConst.調査区分.子牛生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E40", .SHUKEI_DATA1 = "G10:H40", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U26", .SHUKEI_DATA3 = "X24:AG26"}} _
                , {ComConst.調査区分.乳用雄育成牛生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U27", .SHUKEI_DATA3 = "X24:AG27"}} _
                , {ComConst.調査区分.交雑種育成牛生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U27", .SHUKEI_DATA3 = "X24:AG27"}} _
                , {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U27", .SHUKEI_DATA3 = "X24:AG27"}} _
                , {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U27", .SHUKEI_DATA3 = "X24:AG27"}} _
                , {ComConst.調査区分.交雑種肥育牛生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U27", .SHUKEI_DATA3 = "X24:AG27"}} _
                , {ComConst.調査区分.肥育豚生産費統計_個別, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E41", .SHUKEI_DATA1 = "G10:H41", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
                , {ComConst.調査区分.経営分析調査_二条大麦生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_六条大麦生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_はだか麦生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_そば生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_なたね生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_てんさい生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_さとうきび生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L32", .PER3_SHU = "X32", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L34", .PREV_YEAR_3 = "R34", .THIS_YEAR_3_SHU = "X34", .PREV_YEAR_3_SHU = "AD34", .KOBETSU_DATA1 = "D10:E38", .SHUKEI_DATA1 = "G10:H38", .KOBETSU_DATA2 = "L10:P28", .SHUKEI_DATA2 = "X10:AB28", .KOBETSU_DATA3 = "L36:U39", .SHUKEI_DATA3 = "X36:AG39"}} _
                , {ComConst.調査区分.経営分析調査_牛乳生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L20", .PER3_SHU = "X20", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L22", .PREV_YEAR_3 = "R22", .THIS_YEAR_3_SHU = "X22", .PREV_YEAR_3_SHU = "AD22", .KOBETSU_DATA1 = "D10:E40", .SHUKEI_DATA1 = "G10:H40", .KOBETSU_DATA2 = "L10:P16", .SHUKEI_DATA2 = "X10:AB16", .KOBETSU_DATA3 = "L24:U29", .SHUKEI_DATA3 = "X24:AG29"}} _
                , {ComConst.調査区分.経営分析調査_子牛生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E40", .SHUKEI_DATA1 = "G10:H40", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U25", .SHUKEI_DATA3 = "X23:AG25"}} _
                , {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
                , {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
                , {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
                , {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
                , {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E39", .SHUKEI_DATA1 = "G10:H39", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
                , {ComConst.調査区分.経営分析調査_肥育豚生産費, New CellAddress With {.PER1 = "D7", .PER1_SHU = "G7", .PER2 = "L7", .PER2_SHU = "X7", .PER3 = "L19", .PER3_SHU = "X19", .THIS_YEAR_1 = "D9", .PREV_YEAR_1 = "E9", .THIS_YEAR_1_SHU = "G9", .PREV_YEAR_1_SHU = "H9", .THIS_YEAR_2 = "L9", .PREV_YEAR_2 = "P9", .THIS_YEAR_2_SHU = "X9", .PREV_YEAR_2_SHU = "AB9", .THIS_YEAR_3 = "L21", .PREV_YEAR_3 = "R21", .THIS_YEAR_3_SHU = "X21", .PREV_YEAR_3_SHU = "AD21", .KOBETSU_DATA1 = "D10:E41", .SHUKEI_DATA1 = "G10:H41", .KOBETSU_DATA2 = "L10:P15", .SHUKEI_DATA2 = "X10:AB15", .KOBETSU_DATA3 = "L23:U26", .SHUKEI_DATA3 = "X23:AG26"}} _
            }

            Public Const KAISETSU1 As String = "角丸四角形吹き出し 12"
            Public Const KAISETSU2 As String = "角丸四角形吹き出し 15"

        End Class

        Public Class sheet4
            Public Shared PER As String = "C6"
            Public Shared THIS_YEAR As String = "C8"
            Public Shared PREV_YEAR As String = "D8"

            Public Shared TITLE As String = "A9:B37"
            Public Shared KOBETSU_DATA As String = "C9:D37"

            Public Shared DATA_ROW As Integer = 29
        End Class

        Public Class sheet5
            Public Shared PER As String = "C7"
            Public Shared PER_SHU As String = "F7"
            Public Shared THIS_YEAR As String = "C9"
            Public Shared PREV_YEAR As String = "D9"
            Public Shared THIS_YEAR_SHU As String = "F9"
            Public Shared PREV_YEAR_SHU As String = "G9"

            Public Shared TITLE As String = "A10:B38"
            Public Shared KOBETSU_DATA As String = "C10:D38"
            Public Shared SHUKEI_DATA As String = "F10:G38"

            Public Shared DATA_ROW As Integer = 29
        End Class

        Public Class sheet6
            Public Const TITLE As String = "角丸四角形 55"
            Public Const GRAPH As String = "正方形/長方形 68"
            Public Const KAISETSU As String = "角丸四角形吹き出し 50"

            Public Const KOUMOKU_NAME As String = "A63:A70"
            Public Const DATA As String = "E63:F70"

            Public Shared 解説 As New Dictionary(Of String, String) From { _
                {ComConst.調査区分.米生産費統計_個別, "◎　上の図は、米の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.小麦生産費統計_個別, "◎　上の図は、小麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.二条大麦生産費統計_個別, "◎　上の図は、二条大麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.六条大麦生産費統計_個別, "◎　上の図は、六条大麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.はだか麦生産費統計_個別, "◎　上の図は、はだか麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.そば生産費統計_個別, "◎　上の図は、そばの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.大豆生産費統計_個別, "◎　上の図は、大豆の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.原料用かんしょ生産費統計_個別, "◎　上の図は、原料用かんしょの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "◎　上の図は、原料用ばれいしょの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.なたね生産費統計_個別, "◎　上の図は、なたねの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.てんさい生産費統計_個別, "◎　上の図は、てんさいの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.さとうきび生産費統計_個別, "◎　上の図は、さとうきびの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.米生産費統計_組織法人, "◎　上の図は、米の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.小麦生産費統計_組織法人, "◎　上の図は、小麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.大豆生産費統計_組織法人, "◎　上の図は、大豆の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.牛乳生産費統計_個別, "◎　上の図は、生乳の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　乳脂肪分3.5%換算乳量、搾乳牛通年換算頭数及び乳脂肪分については、青い円より外側であれば、{0}より乳量、飼養頭数が多く、乳脂肪分が高いということになります。" & vbLf & "　搾乳牛通年換算１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.子牛生産費統計_個別, "◎　上の図は、子牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　子牛販売頭数及び繁殖雌牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　子牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.乳用雄育成牛生産費統計_個別, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.交雑種育成牛生産費統計_個別, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.交雑種肥育牛生産費統計_個別, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.肥育豚生産費統計_個別, "◎　上の図は、肥育豚の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　肥育豚販売頭数及び肉豚飼養月平均頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　肥育豚１頭当たり全算入生産費などの費用金額と労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_二条大麦生産費, "◎　上の図は、二条大麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_六条大麦生産費, "◎　上の図は、六条大麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_はだか麦生産費, "◎　上の図は、はだか麦の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_そば生産費, "◎　上の図は、そばの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "◎　上の図は、原料用ばれいしょの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_なたね生産費, "◎　上の図は、なたねの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_てんさい生産費, "◎　上の図は、てんさいの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_さとうきび生産費, "◎　上の図は、さとうきびの生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　主産物10a当たり収量、作付面積については、青い円より外側であれば、{0}より単収が多く、作付面積が大きいということになります。" & vbLf & "　10a当たり全算入生産費などの費用金額、資本額や労働時間については、青い円より外側で あれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_牛乳生産費, "◎　上の図は、生乳の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　乳脂肪分3.5%換算乳量、搾乳牛通年換算頭数及び乳脂肪分については、青い円より外側であれば、{0}より乳量、飼養頭数が多く、乳脂肪分が高いということになります。" & vbLf & "　搾乳牛通年換算１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_子牛生産費, "◎　上の図は、子牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　子牛販売頭数及び繁殖雌牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　子牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, "◎　上の図は、育成牛・肥育牛の生産における指標となるデータを、{0}を１００とした場合と貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　育成牛・肥育牛販売頭数及び育成牛・肥育牛飼養頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　育成牛・肥育牛１頭当たり全算入生産費などの費用金額と直接労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"},
                {ComConst.調査区分.経営分析調査_肥育豚生産費, "◎　上の図は、肥育豚の生産における指標となるデータを、{0}を１００とした場合の貴殿の比率をレーダーチャートで示したものです。" & vbLf & "　肥育豚販売頭数及び肉豚飼養月平均頭数については、青い円より外側であれば、{0}より販売頭数、飼養頭数が多いということになります。" & vbLf & "　肥育豚１頭当たり全算入生産費などの費用金額と労働時間については、青い円より外側であれば、{0}より生産に要した金額、時間が多いということになります。"}
            }
        End Class

        Public Class sheet7
            Public Const TITLE1 As String = "正方形/長方形 4"
            Public Const TITLE2 As String = "正方形/長方形 3"
            Public Const SEISANHI As String = "D61:G64"
            Public Const SEISANGAIKYO As String = "B39:F42"


            Public Shared 調査種類_その他 As New Dictionary(Of String, String) From { _
                    {ComConst.調査区分.米生産費統計_個別, "K010930"},
                    {ComConst.調査区分.小麦生産費統計_個別, "K010932"},
                    {ComConst.調査区分.二条大麦生産費統計_個別, "K010932"},
                    {ComConst.調査区分.六条大麦生産費統計_個別, "K010932"},
                    {ComConst.調査区分.はだか麦生産費統計_個別, "K010932"},
                    {ComConst.調査区分.そば生産費統計_個別, "K010932"},
                    {ComConst.調査区分.大豆生産費統計_個別, "K010932"},
                    {ComConst.調査区分.原料用かんしょ生産費統計_個別, "K010932"},
                    {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, "K010932"},
                    {ComConst.調査区分.なたね生産費統計_個別, "K010932"},
                    {ComConst.調査区分.てんさい生産費統計_個別, "K010932"},
                    {ComConst.調査区分.さとうきび生産費統計_個別, "K010932"},
                    {ComConst.調査区分.米生産費統計_組織法人, "K010932"},
                    {ComConst.調査区分.小麦生産費統計_組織法人, "K010932"},
                    {ComConst.調査区分.大豆生産費統計_組織法人, "K010932"},
                    {ComConst.調査区分.牛乳生産費統計_個別, "K010936"},
                    {ComConst.調査区分.子牛生産費統計_個別, "K010935"},
                    {ComConst.調査区分.乳用雄育成牛生産費統計_個別, "K010934"},
                    {ComConst.調査区分.交雑種育成牛生産費統計_個別, "K010934"},
                    {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, "K010934"},
                    {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, "K010934"},
                    {ComConst.調査区分.交雑種肥育牛生産費統計_個別, "K010934"},
                    {ComConst.調査区分.肥育豚生産費統計_個別, "K010938"},
                    {ComConst.調査区分.経営分析調査_二条大麦生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_六条大麦生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_はだか麦生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_そば生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_なたね生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_てんさい生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_さとうきび生産費, "K010932"},
                    {ComConst.調査区分.経営分析調査_牛乳生産費, "K010936"},
                    {ComConst.調査区分.経営分析調査_子牛生産費, "K010935"},
                    {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, "K010934"},
                    {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, "K010934"},
                    {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, "K010934"},
                    {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, "K010934"},
                    {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, "K010934"},
                    {ComConst.調査区分.経営分析調査_肥育豚生産費, "K010938"}
               }


        End Class


        Public Class sheet8
            Public Const KAISETSU1 As String = "角丸四角形吹き出し 5"
            Public Const KAISETSU2 As String = "角丸四角形吹き出し 7"
            Public Const TITLE1 As String = "角丸四角形 2"
            Public Const TITLE2 As String = "角丸四角形 3"

            Public Const HIMOKU As String = "A61:D66"
            Public Const ROUDOU As String = "A70:D75"

            Public Const GRAPH1 As String = "グラフ 1"
            Public Const GRAPH2 As String = "グラフ 4"

            Public Class kaisetsu
                Public 主要費目 As String
                Public 労働時間 As String
            End Class

            Public Shared 項目の解説 As New Dictionary(Of String, kaisetsu) From { _
                            {ComConst.調査区分.米生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、米生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、米の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.小麦生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、小麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、小麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.二条大麦生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、二条大麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、二条大麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.六条大麦生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、六条大麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、六条大麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.はだか麦生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、はだか麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、はだか麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.そば生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、そば生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、そばの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.大豆生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、大豆生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、大豆の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.原料用かんしょ生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、原料用かんしょ生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、原料用かんしょの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、原料用ばれいしょ生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、原料用ばれいしょの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.なたね生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、なたね生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、なたねの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.てんさい生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、てんさい生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、てんさいの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.さとうきび生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、さとうきび生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、さとうきびの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.米生産費統計_組織法人, New kaisetsu With {.主要費目 = "◎　上の図は、米生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、米の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.小麦生産費統計_組織法人, New kaisetsu With {.主要費目 = "◎　上の図は、小麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、小麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.大豆生産費統計_組織法人, New kaisetsu With {.主要費目 = "◎　上の図は、大豆生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、大豆の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.牛乳生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の搾乳牛通年換算１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、生乳の生産にかかる労働時間について、{0}貴殿の搾乳牛通年換算１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.子牛生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の子牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、子牛生産にかかる労働時間について、{0}貴殿の子牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.乳用雄育成牛生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、乳用雄育成牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.交雑種育成牛生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、交雑種育成牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、去勢若齢肥育牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、乳用雄肥育牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.交雑種肥育牛生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、交雑種肥育牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.肥育豚生産費統計_個別, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の肥育豚１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、肥育豚の生産にかかる労働時間について、{0}貴殿の肥育豚１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_二条大麦生産費, New kaisetsu With {.主要費目 = "◎　上の図は、二条大麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、二条大麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_六条大麦生産費, New kaisetsu With {.主要費目 = "◎　上の図は、六条大麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、六条大麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_はだか麦生産費, New kaisetsu With {.主要費目 = "◎　上の図は、はだか麦生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、はだか麦の生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_そば生産費, New kaisetsu With {.主要費目 = "◎　上の図は、そば生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、そばの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, New kaisetsu With {.主要費目 = "◎　上の図は、原料用ばれいしょ生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、原料用ばれいしょの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_なたね生産費, New kaisetsu With {.主要費目 = "◎　上の図は、なたね生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、なたねの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_てんさい生産費, New kaisetsu With {.主要費目 = "◎　上の図は、てんさい生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、てんさいの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_さとうきび生産費, New kaisetsu With {.主要費目 = "◎　上の図は、さとうきび生産費調査の費用のうち主要なものについて、{0}貴殿の10ａ当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、さとうきびの生産にかかる主要な作業について、{0}貴殿の10ａ当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_牛乳生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の搾乳牛通年換算１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、生乳の生産にかかる労働時間について、{0}貴殿の搾乳牛通年換算１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_子牛生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の子牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、子牛生産にかかる労働時間について、{0}貴殿の子牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、乳用雄育成牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、交雑種育成牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、去勢若齢肥育牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、乳用雄肥育牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、交雑種肥育牛の生産にかかる労働時間について、{0}貴殿の育成牛・肥育牛１頭当たりのデータを比較したものです。"}},
                            {ComConst.調査区分.経営分析調査_肥育豚生産費, New kaisetsu With {.主要費目 = "◎　上の図は、生産費の費用のうち主要費目について、{0}貴殿の肥育豚１頭当たりのデータを比較したものです。", .労働時間 = "◎　上の図は、肥育豚の生産にかかる労働時間について、{0}貴殿の肥育豚１頭当たりのデータを比較したものです。"}}
            }

        End Class

        Public Class sheet9
            Public Class Koumoku
                Public 物財費 As String
                Public 労働費 As String
                Public その他 As String
                Public 労働時間 As String

            End Class

            Public Shared 調査種類 As New Dictionary(Of String, Koumoku) From { _
                {ComConst.調査区分.米生産費統計_個別, New Koumoku With {.物財費 = "K010906", .労働費 = "K010919", .その他 = "K010930", .労働時間 = "K010850"}},
                {ComConst.調査区分.小麦生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.二条大麦生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.六条大麦生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.はだか麦生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.そば生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.大豆生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.原料用かんしょ生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.なたね生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.てんさい生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.さとうきび生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010855"}},
                {ComConst.調査区分.米生産費統計_組織法人, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010451"}},
                {ComConst.調査区分.小麦生産費統計_組織法人, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010453"}},
                {ComConst.調査区分.大豆生産費統計_組織法人, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010453"}},
                {ComConst.調査区分.牛乳生産費統計_個別, New Koumoku With {.物財費 = "K010909", .労働費 = "K010925", .その他 = "K010936", .労働時間 = "K020814"}},
                {ComConst.調査区分.子牛生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010924", .その他 = "K010935", .労働時間 = "K020813"}},
                {ComConst.調査区分.乳用雄育成牛生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020813"}},
                {ComConst.調査区分.交雑種育成牛生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020813"}},
                {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020813"}},
                {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020813"}},
                {ComConst.調査区分.交雑種肥育牛生産費統計_個別, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020813"}},
                {ComConst.調査区分.肥育豚生産費統計_個別, New Koumoku With {.物財費 = "K010909", .労働費 = "K010927", .その他 = "K010938", .労働時間 = "K021012"}},
                {ComConst.調査区分.経営分析調査_二条大麦生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_六条大麦生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_はだか麦生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_そば生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_なたね生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_てんさい生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_さとうきび生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010921", .その他 = "K010932", .労働時間 = "K010455"}},
                {ComConst.調査区分.経営分析調査_牛乳生産費, New Koumoku With {.物財費 = "K010909", .労働費 = "K010925", .その他 = "K010936", .労働時間 = "K020413"}},
                {ComConst.調査区分.経営分析調査_子牛生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010924", .その他 = "K010935", .労働時間 = "K020412"}},
                {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020412"}},
                {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020412"}},
                {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020412"}},
                {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020412"}},
                {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, New Koumoku With {.物財費 = "K010908", .労働費 = "K010923", .その他 = "K010934", .労働時間 = "K020412"}},
                {ComConst.調査区分.経営分析調査_肥育豚生産費, New Koumoku With {.物財費 = "K010909", .労働費 = "K010927", .その他 = "K010938", .労働時間 = "K020412"}}
}

            Public Const 全算入生産費 As String = "B62:F65"
            Public Const 労働時間 As String = "B69:F70"

            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From { _
                 {1, "2"}, {2, "1,3"}, {3, "0,2,4"}, {4, "0,1,2,3"}, {5, "0,1,2,3,4"} _
                }

        End Class

        Public Class sheet10
            Public Class Koumoku
                Public 物財費合計 As String
                Public 収量 As String
                Public 規模 As String
            End Class

            Public Const GRAPH As String = "グラフ 3"

            Public Shared 調査種類 As New Dictionary(Of String, Koumoku) From { _
                {ComConst.調査区分.米生産費統計_個別, New Koumoku With {.物財費合計 = "K010906", .規模 = "K011416", .収量 = "K011349"}},
                {ComConst.調査区分.小麦生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.二条大麦生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.六条大麦生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.はだか麦生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.そば生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.大豆生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.原料用かんしょ生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.原料用ばれいしょ生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.なたね生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.てんさい生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.さとうきび生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.米生産費統計_組織法人, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.小麦生産費統計_組織法人, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.大豆生産費統計_組織法人, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.牛乳生産費統計_個別, New Koumoku With {.物財費合計 = "K010909", .規模 = "K011243", .収量 = "K010343"}},
                {ComConst.調査区分.子牛生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011244", .収量 = "K010440"}},
                {ComConst.調査区分.乳用雄育成牛生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.交雑種育成牛生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.乳用雄肥育牛生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.交雑種肥育牛生産費統計_個別, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.肥育豚生産費統計_個別, New Koumoku With {.物財費合計 = "K010909", .規模 = "K011242", .収量 = "K010241"}},
                {ComConst.調査区分.経営分析調査_二条大麦生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_六条大麦生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_はだか麦生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_そば生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_なたね生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_てんさい生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_さとうきび生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011418", .収量 = "K011207"}},
                {ComConst.調査区分.経営分析調査_牛乳生産費, New Koumoku With {.物財費合計 = "K010909", .規模 = "K011243", .収量 = "K010343"}},
                {ComConst.調査区分.経営分析調査_子牛生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011244", .収量 = "K010440"}},
                {ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.経営分析調査_交雑種育成牛生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, New Koumoku With {.物財費合計 = "K010908", .規模 = "K011242", .収量 = "K010441"}},
                {ComConst.調査区分.経営分析調査_肥育豚生産費, New Koumoku With {.物財費合計 = "K010909", .規模 = "K011242", .収量 = "K010241"}}
  }
            Public Const 規模 As String = "A62:F63"
            Public Const 収量 As String = "A65:F66"
            Public Const 物財費 As String = "A70:F76"

            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From { _
                 {1, "3"}, {2, "2,4"}, {3, "1,3,5"}, {4, "1,2,3,4"}, {5, "1,2,3,4,5"} _
                }

        End Class

        Public Class sheet11
            Public Shared お問合せ先 As New Dictionary(Of Integer, String) From {
               {ComConst.還元資料.お問合せ先_明細番号.事務局_拠点, "E46"},
               {ComConst.還元資料.お問合せ先_明細番号.住所, "E48"},
               {ComConst.還元資料.お問合せ先_明細番号.電話番号, "AA50"},
               {ComConst.還元資料.お問合せ先_明細番号.FAX, "AA51"},
               {ComConst.還元資料.お問合せ先_明細番号.担当者名, "Z53"}
             }

        End Class
    End Class

    Public Sub New(pCheckBox As ArrayList, pPrintFlg As Boolean, pOutputFlg As Boolean, pOutputPath As String, pSetting As DataTable,
               pKobetsu As Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)),
               pShukei As Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)),
               pMaster As DataTable, pTaiouMaster As DataTable, pKoumokuMaster As DataTable, pChosaNenKey As ArrayList,
               pSeidoUketoriItem As DataTable) ' REV 001

        ' REV_001↓
        'MyBase.New(pCheckBox, pPrintFlg, pOutputFlg, pOutputPath, pSetting, pKobetsu, pShukei, pMaster, pTaiouMaster, pKoumokuMaster, pChosaNenKey)
        MyBase.New(pCheckBox, pPrintFlg, pOutputFlg, pOutputPath, pSetting, pKobetsu, pShukei, pMaster, pTaiouMaster, pKoumokuMaster, pChosaNenKey, pSeidoUketoriItem)
        ' REV_001↑
    End Sub

    ''' <summary>
    ''' 作成実行
    ''' </summary>
    ''' <param name="pCensusNo">センサス番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)
        Dim xlSheet As Excel.Worksheet = Nothing

        '不要なシートのリスト
        Dim noSheetName As New ArrayList

        '1シート目
        xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
        If CType(msheetList(0), Boolean) = True Then
            CreateSheet1(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '2シート目
        xlSheet = DirectCast(xlSheets.Item(2), Excel.Worksheet)
        If CType(msheetList(1), Boolean) = True Then
            CreateSheet2(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '3シート目
        xlSheet = DirectCast(xlSheets.Item(3), Excel.Worksheet)
        If CType(msheetList(2), Boolean) = True Then
            CreateSheet3(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '4シート目
        xlSheet = DirectCast(xlSheets.Item(4), Excel.Worksheet)
        If CType(msheetList(3), Boolean) = True Then
            CreateSheet4(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '5シート目
        xlSheet = DirectCast(xlSheets.Item(5), Excel.Worksheet)
        If CType(msheetList(4), Boolean) = True Then
            CreateSheet5(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '6シート目
        xlSheet = DirectCast(xlSheets.Item(6), Excel.Worksheet)
        If CType(msheetList(5), Boolean) = True Then
            CreateSheet6(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '7シート目
        xlSheet = DirectCast(xlSheets.Item(7), Excel.Worksheet)
        If CType(msheetList(6), Boolean) = True Then
            CreateSheet7(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '8シート目
        xlSheet = DirectCast(xlSheets.Item(8), Excel.Worksheet)
        If CType(msheetList(7), Boolean) = True Then
            CreateSheet8(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '9シート目
        xlSheet = DirectCast(xlSheets.Item(9), Excel.Worksheet)
        If CType(msheetList(8), Boolean) = True Then
            CreateSheet9(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '10シート目
        xlSheet = DirectCast(xlSheets.Item(10), Excel.Worksheet)
        If CType(msheetList(9), Boolean) = True Then
            CreateSheet10(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '11シート目
        xlSheet = DirectCast(xlSheets.Item(11), Excel.Worksheet)
        If CType(msheetList(10), Boolean) = True Then
            CreateSheet11(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '不要なシートの削除
        If noSheetName.Count <> 0 Then
            For Each sheetName As String In noSheetName
                xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                xlSheet.Delete()
                ReleaseComObject(xlSheet)
                xlSheet = Nothing
            Next
        End If

        'フッタの設定
        Dim firstPageFlg As Boolean = True
        Dim cencusNo As String

        cencusNo = mKobetsu(CStr(mChosanenKey(0))).Item(CENCUS).値
        cencusNo = cencusNo.Substring(0, 2) & Space(1) & cencusNo.Substring(2, 3) & Space(1) & cencusNo.Substring(5, 2) & Space(1) & cencusNo.Substring(7, 3) & _
                    Space(1) & cencusNo.Substring(10, 3) & Space(1) & cencusNo.Substring(13, 3)
        For i As Integer = 1 To xlSheets.Count
            xlSheet = DirectCast(xlSheets.Item(i), Excel.Worksheet)

            xlSheet.PageSetup.RightFooter = cencusNo
            If xlSheet.Name <> "還元１" Then
                If firstPageFlg = True Then
                    xlSheet.PageSetup.FirstPageNumber = 1
                    firstPageFlg = False
                End If
            End If
        Next

    End Sub

    Private Sub CreateSheet1(pXlsheet As Excel.Worksheet)
        Dim range As Excel.Range = Nothing

        Dim yearOptionString As String = String.Empty
        If ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.農産物生産費 Then
            yearOptionString = "産"
        End If

        '和暦の設定
        range = pXlsheet.Range(生産費.sheet1.WAREKI_CELL)
        range.Value = getNengo(mChosanenKey(0).ToString, False) + yearOptionString
        ReleaseComObject(range)
        range = Nothing

        '営農類型の設定
        Dim einou As String = mKobetsu(CStr(mChosanenKey(0))).Item("K000005").値
        range = pXlsheet.Range(生産費.sheet1.KUBUN_CELL)
        range.Value = "【" + ComConst.還元資料.表示名(CommonInfo.Chosakubun) + "生産費】"
        'range.Value = ComConst.営農類型区分.リスト(einou) & "経営"
        ReleaseComObject(range)
        range = Nothing

        '画像の設定
        Dim filePath As String = IniFileInfo.ExcelReportPath() & "\" & FOLDER_NAME & "\" & GetPictureFileName()

        Dim xlShapes As Excel.Shapes
        range = pXlsheet.Range(生産費.sheet1.PICTURE_CELL)
        xlShapes = pXlsheet.Shapes
        Dim xlShape As Excel.Shape

        xlShape = xlShapes.AddPicture(Filename:=filePath, _
           LinkToFile:=Microsoft.Office.Core.MsoTriState.msoFalse, SaveWithDocument:=Microsoft.Office.Core.MsoTriState.msoTrue, _
           Left:=CSng(range.Left), Top:=CSng(range.Top), Width:=391, Height:=261.6)

        '挨拶文の設定
        range = pXlsheet.Range(生産費.sheet1.AISATSU_CELL)
        Dim tmp As String = CStr(range.Value)
        If ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.農産物生産費 Then
            tmp = tmp.Replace("年産", getNengo(mChosanenKey(0).ToString, False) + yearOptionString)
        Else
            tmp = tmp.Replace("年", getNengo(mChosanenKey(0).ToString, False) + yearOptionString)
        End If

        range.Value = tmp

    End Sub

    Private Sub CreateSheet2(xlsheet As Excel.Worksheet)

        Dim cell As 生産費.sheet2.CellAddress = 生産費.sheet2.CellList(CommonInfo.Chosakubun)
        Dim define As 生産費.sheet2.Define = 生産費.sheet2.DefineList(CommonInfo.Chosakubun)

        '表頭
        SetKangenShiryoYear(xlsheet, cell.THIS_YEAR_1, cell.PREV_YEAR_1)
        SetKangenShiryoYear(xlsheet, cell.THIS_YEAR_2, cell.PREV_YEAR_2)
        SetKangenShiryoYear(xlsheet, cell.THIS_YEAR_3, cell.PREV_YEAR_3)

        '固定項目
        SetKangenShiryoFixed(xlsheet, 2, 1, define.DATA1_ROW, cell.KOBETSU_DATA1)
        SetKangenShiryoFixed(xlsheet, 2, 2, define.DATA2_ROW, cell.KOBETSU_DATA2, String.Empty, 3)
        SetKangenShiryoFixed2Col(xlsheet, 2, 3, define.DATA3_ROW, cell.KOBETSU_DATA3, String.Empty, 2)

        '用語の解説
        Dim text As String = xlsheet.Shapes.Item(生産費.sheet2.KAISETSU1).TextEffect.Text
        text = text.Replace("ＮＮＮＮＮＮＮＮ", ComConst.還元資料.表示名(CommonInfo.Chosakubun))
        If ComConst.還元資料.単位当たり表示.ContainsKey(CommonInfo.Chosakubun) Then
            '単位当たり表示を行う調査区分のみ置き換え
            text = text.Replace("999NN", ComConst.還元資料.単位当たり表示(CommonInfo.Chosakubun))
        End If
        xlsheet.Shapes.Item(生産費.sheet2.KAISETSU1).TextEffect.Text = text
    End Sub

    Private Sub CreateSheet3(xlSheet As Excel.Worksheet)

        Dim cell As 生産費.sheet3.CellAddress = 生産費.sheet3.リスト(CommonInfo.Chosakubun)
        Dim define As 生産費.sheet2.Define = 生産費.sheet2.DefineList(CommonInfo.Chosakubun)  ' 行数は還元2と同様なので流用

        '表頭
        ReplaceKangenShiryoHeikin(xlSheet, cell.PER1_SHU)
        ReplaceKangenShiryoHeikin(xlSheet, cell.PER2_SHU)
        ReplaceKangenShiryoHeikin(xlSheet, cell.PER3_SHU)

        SetKangenShiryoYear(xlSheet, cell.THIS_YEAR_1, cell.PREV_YEAR_1)
        SetKangenShiryoYear(xlSheet, cell.THIS_YEAR_1_SHU, cell.PREV_YEAR_1_SHU)
        SetKangenShiryoYear(xlSheet, cell.THIS_YEAR_2, cell.PREV_YEAR_2)
        SetKangenShiryoYear(xlSheet, cell.THIS_YEAR_2_SHU, cell.PREV_YEAR_2_SHU)
        SetKangenShiryoYear(xlSheet, cell.THIS_YEAR_3, cell.PREV_YEAR_3)
        SetKangenShiryoYear(xlSheet, cell.THIS_YEAR_3_SHU, cell.PREV_YEAR_3_SHU)

        '固定項目
        SetKangenShiryoFixed(xlSheet, 2, 1, define.DATA1_ROW, cell.KOBETSU_DATA1, cell.SHUKEI_DATA1)
        SetKangenShiryoFixed(xlSheet, 2, 2, define.DATA2_ROW, cell.KOBETSU_DATA2, cell.SHUKEI_DATA2, 3, 3)
        SetKangenShiryoFixed2Col(xlSheet, 2, 3, define.DATA3_ROW, cell.KOBETSU_DATA3, cell.SHUKEI_DATA3, 2, 2)

        '用語の解説
        Dim text As String = xlSheet.Shapes.Item(生産費.sheet3.KAISETSU1).TextEffect.Text
        text = text.Replace("ＮＮＮＮＮＮＮＮ", ComConst.還元資料.表示名(CommonInfo.Chosakubun))
        If ComConst.還元資料.単位当たり表示.ContainsKey(CommonInfo.Chosakubun) Then
            '単位当たり表示を行う調査区分のみ置き換え
            text = text.Replace("999NN", ComConst.還元資料.単位当たり表示(CommonInfo.Chosakubun))
        End If
        text = text.Replace("ＮＮＮ平均", getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
        xlSheet.Shapes.Item(生産費.sheet3.KAISETSU1).TextEffect.Text = text
    End Sub

    Private Sub CreateSheet4(xlSheet As Excel.Worksheet)
        '統計表　表頭
        SetKangenShiryoYear(xlSheet, 生産費.sheet4.THIS_YEAR, 生産費.sheet4.PREV_YEAR)

        '統計表 任意項目
        SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_生産費.還元4_統計表, 生産費.sheet4.DATA_ROW, 生産費.sheet4.TITLE, 生産費.sheet4.KOBETSU_DATA)

    End Sub

    Private Sub CreateSheet5(xlSheet As Excel.Worksheet)

        Dim range As Excel.Range = Nothing

        '統計表　表頭
        ReplaceKangenShiryoHeikin(xlSheet, 生産費.sheet5.PER_SHU)
        SetKangenShiryoYear(xlSheet, 生産費.sheet5.THIS_YEAR, 生産費.sheet5.PREV_YEAR)
        SetKangenShiryoYear(xlSheet, 生産費.sheet5.THIS_YEAR_SHU, 生産費.sheet5.PREV_YEAR_SHU)

        '統計表 任意項目
        SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_生産費.還元4_統計表, 生産費.sheet5.DATA_ROW,
                               生産費.sheet5.TITLE, 生産費.sheet5.KOBETSU_DATA, 生産費.sheet5.SHUKEI_DATA)

    End Sub

    Private Sub CreateSheet6(pXlSheet As Excel.Worksheet)
        Dim range As Excel.Range = Nothing
        Dim KoumokuName(7, 0) As Object
        Dim Data(7, 1) As Object
        Dim i As Integer = 0
        Dim d1 As DataRow()

        d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_生産費.還元6_レーダーチャートの表示項目 & "'")

        For Each d2 As DataRow In d1
            If Not String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                KoumokuName(i, 0) = getKoumokuName(d2.Item("設定値").ToString)
                Data(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("設定値").ToString))
                Data(i, 1) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
            End If

            i = i + 1
        Next

        range = pXlSheet.Range(生産費.sheet6.KOUMOKU_NAME)
        range.Value = KoumokuName
        ReleaseComObject(range)
        range = Nothing

        range = pXlSheet.Range(生産費.sheet6.DATA)
        range.Value = Data
        ReleaseComObject(range)
        range = Nothing

        pXlSheet.Shapes.Item(生産費.sheet6.TITLE).TextEffect.Text = pXlSheet.Shapes.Item(生産費.sheet6.TITLE).TextEffect.Text.Replace("年", getNengo(mChosanenKey(0).ToString, False))
        If getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１) <> String.Empty Then
            pXlSheet.Shapes.Item(生産費.sheet6.TITLE).TextEffect.Text = pXlSheet.Shapes.Item(生産費.sheet6.TITLE).TextEffect.Text.Replace("平均", getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
            pXlSheet.Shapes.Item(生産費.sheet6.GRAPH).TextEffect.Text = pXlSheet.Shapes.Item(生産費.sheet6.GRAPH).TextEffect.Text.Replace("平均", getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))

        End If
        pXlSheet.Shapes.Item(生産費.sheet6.KAISETSU).TextEffect.Text = 生産費.sheet6.解説(CommonInfo.Chosakubun).Replace("{0}", getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
    End Sub

    Private Sub CreateSheet7(xlSheet As Excel.Worksheet)
        Dim range As Excel.Range = Nothing
        Dim seisanhi(3, 3) As Object
        Dim seisangaikyo(3, 4) As Object
        Dim heikin1 As String = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
        Dim heikin2 As String = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２)
        Dim heikin3 As String = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３)
        Dim sonota1 As Decimal
        Dim sonota2 As Decimal
        Dim sonota3 As Decimal
        Dim sonota4 As Decimal
        Dim i As Integer = 0
        Dim d1() As DataRow

        seisanhi(0, 0) = "貴殿"
        seisanhi(0, 1) = heikin1
        seisanhi(0, 2) = heikin2
        seisanhi(0, 3) = heikin3

        d1 = mMaster.Select("シート番号 = 7 AND 配列番号 = 1")

        sonota1 = changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(生産費.sheet7.調査種類_その他(CommonInfo.Chosakubun))))
        sonota2 = changeNum(ShukeiDataHenkan(生産費.sheet7.調査種類_その他(CommonInfo.Chosakubun), ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
        sonota3 = changeNum(ShukeiDataHenkan(生産費.sheet7.調査種類_その他(CommonInfo.Chosakubun), ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２))
        sonota4 = changeNum(ShukeiDataHenkan(生産費.sheet7.調査種類_その他(CommonInfo.Chosakubun), ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３))

        i = 2
        For Each d2 As DataRow In d1
            seisanhi(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("個別結果表項番").ToString))
            sonota1 = sonota1 - changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("個別結果表項番").ToString)))

            seisanhi(i, 1) = IIf(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１) = "0", _
                                    String.Empty, ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
            sonota2 = sonota2 - changeNum(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))

            seisanhi(i, 2) = IIf(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２) = "0", _
                                    String.Empty, ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２))
            sonota3 = sonota3 - changeNum(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２))

            seisanhi(i, 3) = IIf(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３) = "0", _
                        String.Empty, ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３))
            sonota4 = sonota4 - changeNum(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３))

            i = i + 1
        Next

        seisanhi(1, 0) = IIf(sonota1 = 0, String.Empty, sonota1)

        seisanhi(1, 1) = IIf(sonota2 = 0, String.Empty, sonota2)

        seisanhi(1, 2) = IIf(sonota3 = 0, String.Empty, sonota3)

        seisanhi(1, 3) = IIf(sonota4 = 0, String.Empty, sonota4)


        seisangaikyo(0, 1) = "貴殿"
        seisangaikyo(0, 2) = heikin1
        seisangaikyo(0, 3) = heikin2
        seisangaikyo(0, 4) = heikin3

        d1 = mMaster.Select("シート番号 = 7 AND 配列番号 = 2")

        i = 1
        For Each d2 As DataRow In d1
            If Not String.IsNullOrEmpty(d2.Item("個別結果表項番").ToString) Then
                seisangaikyo(i, 0) = getKoumokuName(d2.Item("個別結果表項番").ToString) & "(" & getKoumokuName(d2.Item("個別結果表項番").ToString, 1) & ")"
                seisangaikyo(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("個別結果表項番").ToString))
                seisangaikyo(i, 2) = IIf(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１) = "0", _
                                         String.Empty, ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
                seisangaikyo(i, 3) = IIf(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２) = "0", _
                                         String.Empty, ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２))
                seisangaikyo(i, 4) = IIf(ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３) = "0", _
                                  String.Empty, ShukeiDataHenkan(d2.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_３))
            End If
            i = i + 1
        Next


        range = xlSheet.Range(生産費.sheet7.SEISANHI)
        range.Value = seisanhi
        ReleaseComObject(range)
        range = Nothing

        range = xlSheet.Range(生産費.sheet7.SEISANGAIKYO)
        range.Value = seisangaikyo
        ReleaseComObject(range)
        range = Nothing
        If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
            xlSheet.Shapes.Item(生産費.sheet7.TITLE1).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet7.TITLE1).TextEffect.Text.Replace("（年）", "(" & getNengo(mChosanenKey(0).ToString, False) & ")")
        Else
            xlSheet.Shapes.Item(生産費.sheet7.TITLE1).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet7.TITLE1).TextEffect.Text.Replace("年", getNengo(mChosanenKey(0).ToString, False))
        End If

        xlSheet.Shapes.Item(生産費.sheet7.TITLE2).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet7.TITLE2).TextEffect.Text.Replace("年", getNengo(mChosanenKey(0).ToString, False))

    End Sub

    Private Sub CreateSheet8(xlSheet As Excel.Worksheet)
        Dim range As Excel.Range = Nothing
        Dim shuyouhimoku(5, 3) As Object
        Dim roudoujikan(5, 3) As Object
        Dim sheet8 As 生産費.sheet8.kaisetsu = 生産費.sheet8.項目の解説(CommonInfo.Chosakubun)
        Dim d1() As DataRow
        Dim i As Integer
        Dim heikin1 As String = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
        Dim heikin2 As String = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２)

        shuyouhimoku(0, 1) = "貴殿"
        shuyouhimoku(0, 2) = heikin1
        shuyouhimoku(0, 3) = heikin2

        d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.還元資料項目_生産費.還元8_主要項目)

        i = 1
        If d1.Count <> 0 Then
            For Each d2 As DataRow In d1
                If Not String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                    shuyouhimoku(i, 0) = changeSplit(getKoumokuName(d2.Item("設定値").ToString), 7)
                    shuyouhimoku(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("設定値").ToString))
                    shuyouhimoku(i, 2) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    shuyouhimoku(i, 3) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２)

                    i = i + 1
                End If
            Next
        End If

        roudoujikan(0, 1) = "貴殿"
        roudoujikan(0, 2) = heikin1
        roudoujikan(0, 3) = heikin2

        d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.還元資料項目_生産費.還元8_主要作業)
        i = 1

        If d1.Count <> 0 Then
            For Each d2 As DataRow In d1
                If Not String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                    roudoujikan(i, 0) = changeSplit(getKoumokuName(d2.Item("設定値").ToString), 7)
                    roudoujikan(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("設定値").ToString))
                    roudoujikan(i, 2) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    roudoujikan(i, 3) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_２)
                    i = i + 1
                End If
            Next
        End If

        range = xlSheet.Range(生産費.sheet8.HIMOKU)
        range.Value = shuyouhimoku
        ReleaseComObject(range)
        range = Nothing
        Call setGraphSetting(xlSheet, 生産費.sheet8.GRAPH1)

        range = xlSheet.Range(生産費.sheet8.ROUDOU)
        range.Value = roudoujikan
        ReleaseComObject(range)
        range = Nothing
        Call setGraphSetting(xlSheet, 生産費.sheet8.GRAPH2)

        Dim temp As String = String.Empty
        If Not String.IsNullOrEmpty(heikin1) And Not String.IsNullOrEmpty(heikin2) Then
            temp = heikin1 & "、" & heikin2 & "と"
        ElseIf String.IsNullOrEmpty(heikin1) And String.IsNullOrEmpty(heikin2) Then
            temp = String.Empty
        ElseIf String.IsNullOrEmpty(heikin1) OrElse String.IsNullOrEmpty(heikin2) Then
            temp = heikin1 & heikin2 & "と"
        End If


        xlSheet.Shapes.Item(生産費.sheet8.KAISETSU1).TextEffect.Text = 生産費.sheet8.項目の解説(CommonInfo.Chosakubun).主要費目.Replace("{0}", temp)
        xlSheet.Shapes.Item(生産費.sheet8.KAISETSU2).TextEffect.Text = 生産費.sheet8.項目の解説(CommonInfo.Chosakubun).労働時間.Replace("{0}", temp)
        If CommonInfo.Chosakubun = ComConst.調査区分.牛乳生産費統計_個別 OrElse CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_牛乳生産費 Then
            xlSheet.Shapes.Item(生産費.sheet8.TITLE1).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet8.TITLE1).TextEffect.Text.Replace("年・", getNengo(mChosanenKey(0).ToString, False) & "・")
            xlSheet.Shapes.Item(生産費.sheet8.TITLE2).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet8.TITLE2).TextEffect.Text.Replace("年・", getNengo(mChosanenKey(0).ToString, False) & "・")
        Else
            xlSheet.Shapes.Item(生産費.sheet8.TITLE1).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet8.TITLE1).TextEffect.Text.Replace("年", getNengo(mChosanenKey(0).ToString, False))
            xlSheet.Shapes.Item(生産費.sheet8.TITLE2).TextEffect.Text = xlSheet.Shapes.Item(生産費.sheet8.TITLE2).TextEffect.Text.Replace("年", getNengo(mChosanenKey(0).ToString, False))
        End If


    End Sub

    Private Sub CreateSheet9(xlSheet As Excel.Worksheet)
        Dim Range As Excel.Range = Nothing
        Dim seisanhi(3, 4) As Object
        Dim jikan(1, 4) As Object
        Dim start() As String
        Dim sheet9 As 生産費.sheet9.Koumoku = 生産費.sheet9.調査種類(CommonInfo.Chosakubun)

        start = 生産費.sheet9.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)

        Dim yearOptionString As String = String.Empty
        If ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.農産物生産費 Then
            yearOptionString = "年産"
        Else
            yearOptionString = "年"
        End If

        For i As Integer = 1 To start.Count
            seisanhi(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True) & yearOptionString
            seisanhi(1, CInt(start(i - 1))) = CStr(changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet9.その他))) - _
                                                changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet9.物財費))) - _
                                                changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet9.労働費))))
            If CStr(seisanhi(1, CInt(start(i - 1)))) = "0" Then
                seisanhi(1, CInt(start(i - 1))) = Nothing
            End If
            seisanhi(2, CInt(start(i - 1))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet9.労働費))
            seisanhi(3, CInt(start(i - 1))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet9.物財費))

            jikan(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True) & yearOptionString
            jikan(1, CInt(start(i - 1))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet9.労働時間))
        Next

        Range = xlSheet.Range(生産費.sheet9.全算入生産費)
        Range.Value = seisanhi
        ReleaseComObject(Range)
        Range = Nothing

        Range = xlSheet.Range(生産費.sheet9.労働時間)
        Range.Value = jikan
        ReleaseComObject(Range)
        Range = Nothing
    End Sub

    Private Sub CreateSheet10(xlSheet As Excel.Worksheet)
        Dim Range As Excel.Range = Nothing
        Dim kibo(1, 5) As Object
        Dim suryou(1, 5) As Object
        Dim butsuzaihi(6, 5) As Object
        Dim start() As String
        Dim sheet10 As 生産費.sheet10.Koumoku = 生産費.sheet10.調査種類(CommonInfo.Chosakubun)

        start = 生産費.sheet10.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)

        Dim yearOptionString As String = String.Empty
        If ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.農産物生産費 Then
            yearOptionString = "年産"
        Else
            yearOptionString = "年"
        End If

        For i As Integer = 1 To start.Count
            kibo(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True) & yearOptionString
            kibo(1, CInt(start(i - 1))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet10.規模))
            suryou(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True) & yearOptionString
            suryou(1, CInt(start(i - 1))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet10.収量))
        Next

        Range = xlSheet.Range(生産費.sheet10.規模)
        Range.Value = kibo
        ReleaseComObject(Range)
        Range = Nothing

        Range = xlSheet.Range(生産費.sheet10.収量)
        Range.Value = suryou
        ReleaseComObject(Range)
        Range = Nothing

        Dim d1() As DataRow


        d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_生産費.還元10_物財費内訳 & "'", "明細番号 DESC")

        butsuzaihi(1, 0) = "その他"

        Dim sum As Decimal
        Dim sonota As Decimal

        For i = 1 To start.Count
            sum = 0
            sonota = changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(sheet10.物財費合計)))

            butsuzaihi(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True) & yearOptionString

            For j = 0 To d1.Count - 1
                If Not String.IsNullOrEmpty((d1(j).Item("設定値").ToString)) Then
                    butsuzaihi(2 + j, 0) = getKoumokuName(d1(j).Item("設定値").ToString)
                    butsuzaihi(2 + j, CInt(start(i - 1))) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(d1(j).Item("設定値").ToString))
                    sonota = sonota - changeNum(ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(CStr(d1(j).Item("設定値")))))
                End If
            Next
            If sonota = 0 Then
                butsuzaihi(1, CInt(start(i - 1))) = Nothing
            Else
                butsuzaihi(1, CInt(start(i - 1))) = sonota
            End If

        Next

        Range = xlSheet.Range(生産費.sheet10.物財費)
        Range.Value = butsuzaihi
        ReleaseComObject(Range)
        Range = Nothing

        setGraphSetting(xlSheet, 生産費.sheet10.GRAPH)
    End Sub

    Private Sub CreateSheet11(pXlSheet As Excel.Worksheet)
        'この資料等についての連絡先

        Dim d1 As DataRow()
        Dim range As Excel.Range
        d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.お問合せ先)

        If d1.Count >= 6 Then
            range = pXlSheet.Range(生産費.sheet11.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.事務局_拠点))
            range.Value = d1(0).Item("設定値").ToString + " " + d1(1).Item("設定値").ToString
            ReleaseComObject(range)
            range = Nothing

            range = pXlSheet.Range(生産費.sheet11.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.住所))
            range.Value = d1(5).Item("設定値")
            ReleaseComObject(range)
            range = Nothing

            range = pXlSheet.Range(生産費.sheet11.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.電話番号))
            range.Value = d1(3).Item("設定値")
            ReleaseComObject(range)
            range = Nothing

            range = pXlSheet.Range(生産費.sheet11.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.FAX))
            range.Value = d1(4).Item("設定値")
            ReleaseComObject(range)
            range = Nothing

            range = pXlSheet.Range(生産費.sheet11.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.担当者名))
            range.Value = d1(2).Item("設定値")
            ReleaseComObject(range)
            range = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 画像ファイル名取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPictureFileName() As String
        If CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_個別 Or
            CommonInfo.Chosakubun = ComConst.調査区分.米生産費統計_組織法人 Then
            Return "01_生産費_米生産費.png"
        ElseIf ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.農産物生産費 Then
            Return "02_生産費_畑作生産費.png"
        ElseIf ComConst.調査区分.リスト(CommonInfo.Chosakubun).区分２ = ComConst.区分２.畜産物生産費 Then
            Return "03_生産費_畜産生産費.png"
        Else
            Return String.Empty
        End If
    End Function

    ''' <summary>
    ''' 平均の文字列置換
    ''' </summary>
    ''' <param name="xlSheet"></param>
    ''' <param name="cell"></param>
    ''' <remarks></remarks>
    Private Sub ReplaceKangenShiryoHeikin(xlSheet As Excel.Worksheet, cell As String)
        Dim range As Excel.Range = Nothing
        Dim HeikinChi As String = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
        If String.IsNullOrEmpty(HeikinChi) Then
            HeikinChi = getHeikinchishurui(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
        End If

        range = xlSheet.Range(cell)
        range.Value = range.Value.ToString.Replace("平均", HeikinChi)
        ReleaseComObject(range)
        range = Nothing

    End Sub


    ''' <summary>
    ''' 還元資料固定項目の設定
    ''' </summary>
    ''' <param name="xlSheet"></param>
    ''' <param name="sheetNo"></param>
    ''' <param name="arrayNo"></param>
    ''' <param name="cell"></param>
    ''' <param name="shukeiCell"></param>
    ''' <param name="cellSpace"></param>
    ''' <param name="shukeiCellSpace"></param>
    ''' <remarks></remarks>
    Protected Sub SetKangenShiryoFixed2Col(xlSheet As Excel.Worksheet, sheetNo As Integer, arrayNo As Integer, rowCount As Integer, _
                                     cell As String, Optional shukeiCell As String = "", _
                                     Optional cellSpace As UInteger = 0, Optional shukeiCellSpace As UInteger = 0)
        Dim d1 As DataRow()
        Dim dataArray(rowCount - 1, CInt(1 + cellSpace) * 3) As Object
        Dim dataShukeiArray(rowCount - 1, CInt(1 + shukeiCellSpace) * 3) As Object

        d1 = mMaster.Select("シート番号 = " + sheetNo.ToString + " AND 配列番号 = " + arrayNo.ToString)

        For Each d2 As DataRow In d1

            Dim kobetsuNo As String = d2.Item("個別結果表項番").ToString

            If mKobetsu(CStr(mChosanenKey(0))).ContainsKey(d2.Item("個別結果表項番").ToString) Then

                ' 個別 当年
                dataArray(CInt(d2.Item("ROW")), CInt(CInt(d2.Item("COL")))) = mKobetsu(CStr(mChosanenKey(0))).Item(kobetsuNo).値

                ' 集計
                If String.IsNullOrEmpty(shukeiCell) = False Then
                    Dim d3 As DataRow()
                    Dim seisanHeikin As String = "0"
                    Dim shukeiNo As String = ""

                    d3 = mTaiou.Select("個別結果表項番 = '" + kobetsuNo + "'")

                    '対応する集計結果表を設定
                    GetTaiouShukeiNo(kobetsuNo, seisanHeikin, shukeiNo)

                    ' 集計 当年
                    dataShukeiArray(CInt(d2.Item("ROW")), CInt(d2.Item("COL"))) = _
                        ShukeiDataHenkan(kobetsuNo, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)

                    ' 集計 前年
                    dataShukeiArray(CInt(d2.Item("ROW")), CInt(CInt(d2.Item("COL")) + (1 + shukeiCellSpace) * 2)) = _
                         ShukeiDataHenkan(kobetsuNo, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                End If
            End If

            Dim zennen As Boolean = False
            If mChosanenKey.Count <> 1 Then
                zennen = True
            End If

            ' 個別 前年
            If zennen Then
                If mKobetsu(CStr(mChosanenKey(1))).ContainsKey(kobetsuNo) Then
                    dataArray(CInt(d2.Item("ROW")), CInt((CInt(d2.Item("COL")) + (1 + cellSpace) * 2))) = mKobetsu(CStr(mChosanenKey(1))).Item(kobetsuNo).値
                    ' 汎用的に3列以上も対応する場合は、* 2を変数にする
                End If
            End If
        Next

        '合計値を設定する必要がある項目
        If String.IsNullOrEmpty(shukeiCell) = True Then
            SetKangenShiryoFixedAdditional2Col(sheetNo, arrayNo, dataArray, Nothing, cellSpace)
        Else
            SetKangenShiryoFixedAdditional2Col(sheetNo, arrayNo, dataArray, dataShukeiArray, cellSpace, shukeiCellSpace)
        End If

        Dim range As Excel.Range = Nothing

        Try
            ' 個別結果表 設定値
            range = xlSheet.Range(cell)
            range.Value = dataArray
            ReleaseComObject(range)
            range = Nothing

            If String.IsNullOrEmpty(shukeiCell) = False Then
                ' 集計結果表 設定値
                range = xlSheet.Range(shukeiCell)
                range.Value = dataShukeiArray
                ReleaseComObject(range)
                range = Nothing
            End If
        Finally
            ReleaseComObject(range)
            range = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' 還元資料固定項目の設定(３　主産物・副産物用)
    ''' </summary>
    ''' <param name="sheetNo"></param>
    ''' <param name="arrayNo"></param>
    ''' <param name="kobetsuArray"></param>
    ''' <param name="shukeiArray"></param>
    ''' <param name="cellSpace"></param>
    ''' <param name="shukeiCellSpace"></param>
    ''' <remarks></remarks>
    Protected Sub SetKangenShiryoFixedAdditional2Col(sheetNo As Integer, arrayNo As Integer, ByRef kobetsuArray(,) As Object, ByRef shukeiArray(,) As Object _
                                                    , Optional cellSpace As UInteger = 0, Optional shukeiCellSpace As UInteger = 0)
        If CommonInfo.Chosakubun = ComConst.調査区分.肥育豚生産費統計_個別 OrElse CommonInfo.Chosakubun = ComConst.調査区分.経営分析調査_肥育豚生産費 Then
            '還元資料固定項目定義マスタで保持していない項目の設定

            Dim dcSumInfo As Dictionary(Of String, List(Of 生産費.sheet2.sumInfo2)) = Me.GetKangenShiryoSumInfo2()
            Dim listSumInfo As List(Of 生産費.sheet2.sumInfo2) = dcSumInfo(CommonInfo.Chosakubun)

            Dim kobetsuValue As Decimal = 0

            Dim zennen As Boolean = False
            If mChosanenKey.Count <> 1 Then
                zennen = True
            End If

            For Each sumInfo As 生産費.sheet2.sumInfo2 In listSumInfo
                ' 個別 当年 
                kobetsuArray(sumInfo.ROW, sumInfo.COL) = AddKobetsuValue(
                    mKobetsu(CStr(mChosanenKey(0))).Item(sumInfo.KOBETSU_KOBAN1).値,
                    mKobetsu(CStr(mChosanenKey(0))).Item(sumInfo.KOBETSU_KOBAN2).値,
                    mKobetsu(CStr(mChosanenKey(0))).Item(sumInfo.KOBETSU_KOBAN3).値,
                    mKobetsu(CStr(mChosanenKey(0))).Item(sumInfo.KOBETSU_KOBAN4).値)

                ' 個別 前年
                If zennen Then
                    kobetsuArray(sumInfo.ROW, CInt(sumInfo.COL + (1 + cellSpace) * 2)) = AddKobetsuValue(
                        mKobetsu(CStr(mChosanenKey(1))).Item(sumInfo.KOBETSU_KOBAN1).値,
                        mKobetsu(CStr(mChosanenKey(1))).Item(sumInfo.KOBETSU_KOBAN2).値,
                        mKobetsu(CStr(mChosanenKey(1))).Item(sumInfo.KOBETSU_KOBAN3).値,
                        mKobetsu(CStr(mChosanenKey(1))).Item(sumInfo.KOBETSU_KOBAN4).値)
                End If

                ' 集計
                If Not shukeiArray Is Nothing Then
                    '当年
                    shukeiArray(sumInfo.ROW, sumInfo.COL) = AddKobetsuValue(
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN1, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１),
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN2, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１),
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN3, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１),
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN4, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
                    '前年
                    shukeiArray(sumInfo.ROW, CInt(sumInfo.COL + (1 + shukeiCellSpace) * 2)) = AddKobetsuValue(
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN1, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年),
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN2, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年),
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN3, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年),
                        ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN4, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年))
                End If

            Next
        End If

    End Sub

    ''' <summary>
    ''' 個別結果表２項目合計用の定義取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKangenShiryoSumInfo() As Dictionary(Of String, List(Of 生産費.sheet2.SumInfo))
        Dim SumInfoList As New Dictionary(Of String, List(Of 生産費.sheet2.SumInfo))
        Dim list As New List(Of 生産費.sheet2.SumInfo)

        '米生産費（個別）
        list.Clear()
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 21, .KOBETSU_KOBAN1 = "K010925", .KOBETSU_KOBAN2 = "K010926"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010928", .KOBETSU_KOBAN2 = "K010929"})
        SumInfoList.Add(ComConst.調査区分.米生産費統計_個別, list)
        '麦類・大豆・そば・なたね・畑作物生産費（個別）
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 21, .KOBETSU_KOBAN1 = "K010927", .KOBETSU_KOBAN2 = "K010928"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010930", .KOBETSU_KOBAN2 = "K010931"})
        SumInfoList.Add(ComConst.調査区分.小麦生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.二条大麦生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.六条大麦生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.はだか麦生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.そば生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.大豆生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.原料用かんしょ生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.原料用ばれいしょ生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.なたね生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.てんさい生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.さとうきび生産費統計_個別, list)
        '牛乳生産費（個別）
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010931", .KOBETSU_KOBAN2 = "K010932"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 25, .KOBETSU_KOBAN1 = "K010934", .KOBETSU_KOBAN2 = "K010935"})
        SumInfoList.Add(ComConst.調査区分.牛乳生産費統計_個別, list)
        '子牛生産費(個別)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010930", .KOBETSU_KOBAN2 = "K010931"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 25, .KOBETSU_KOBAN1 = "K010933", .KOBETSU_KOBAN2 = "K010934"})
        SumInfoList.Add(ComConst.調査区分.子牛生産費統計_個別, list)
        '育成牛・肥育牛生産費(個別)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 22, .KOBETSU_KOBAN1 = "K010929", .KOBETSU_KOBAN2 = "K010930"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 24, .KOBETSU_KOBAN1 = "K010932", .KOBETSU_KOBAN2 = "K010933"})
        SumInfoList.Add(ComConst.調査区分.乳用雄育成牛生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.交雑種育成牛生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.去勢若齢肥育牛生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.乳用雄肥育牛生産費統計_個別, list)
        SumInfoList.Add(ComConst.調査区分.交雑種肥育牛生産費統計_個別, list)
        '肥育豚生産費(個別)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 24, .KOBETSU_KOBAN1 = "K010933", .KOBETSU_KOBAN2 = "K010934"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 26, .KOBETSU_KOBAN1 = "K010936", .KOBETSU_KOBAN2 = "K010937"})
        SumInfoList.Add(ComConst.調査区分.肥育豚生産費統計_個別, list)

        '米生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 21, .KOBETSU_KOBAN1 = "K010927", .KOBETSU_KOBAN2 = "K010928"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010930", .KOBETSU_KOBAN2 = "K010931"})
        SumInfoList.Add(ComConst.調査区分.米生産費統計_組織法人, list)
        '小麦・大豆生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 21, .KOBETSU_KOBAN1 = "K010927", .KOBETSU_KOBAN2 = "K010928"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010930", .KOBETSU_KOBAN2 = "K010931"})
        SumInfoList.Add(ComConst.調査区分.小麦生産費統計_組織法人, list)
        SumInfoList.Add(ComConst.調査区分.大豆生産費統計_組織法人, list)
        '麦類・そば・なたね・畑作物生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 21, .KOBETSU_KOBAN1 = "K010927", .KOBETSU_KOBAN2 = "K010928"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010930", .KOBETSU_KOBAN2 = "K010931"})
        SumInfoList.Add(ComConst.調査区分.経営分析調査_二条大麦生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_六条大麦生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_はだか麦生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_そば生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_原料用ばれいしょ生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_なたね生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_てんさい生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_さとうきび生産費, list)
        '牛乳生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010931", .KOBETSU_KOBAN2 = "K010932"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 25, .KOBETSU_KOBAN1 = "K010934", .KOBETSU_KOBAN2 = "K010935"})
        SumInfoList.Add(ComConst.調査区分.経営分析調査_牛乳生産費, list)
        '子牛生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 23, .KOBETSU_KOBAN1 = "K010930", .KOBETSU_KOBAN2 = "K010931"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 25, .KOBETSU_KOBAN1 = "K010933", .KOBETSU_KOBAN2 = "K010934"})
        SumInfoList.Add(ComConst.調査区分.経営分析調査_子牛生産費, list)
        '育成牛・肥育牛生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 22, .KOBETSU_KOBAN1 = "K010929", .KOBETSU_KOBAN2 = "K010930"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 24, .KOBETSU_KOBAN1 = "K010932", .KOBETSU_KOBAN2 = "K010933"})
        SumInfoList.Add(ComConst.調査区分.経営分析調査_乳用雄育成牛生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_交雑種育成牛生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_去勢若齢肥育牛生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_乳用雄肥育牛生産費, list)
        SumInfoList.Add(ComConst.調査区分.経営分析調査_交雑種肥育牛生産費, list)
        '肥育豚生産費(組織法人)
        list = New List(Of 生産費.sheet2.SumInfo)
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 24, .KOBETSU_KOBAN1 = "K010933", .KOBETSU_KOBAN2 = "K010934"})
        list.Add(New 生産費.sheet2.SumInfo With {.ROW = 26, .KOBETSU_KOBAN1 = "K010936", .KOBETSU_KOBAN2 = "K010937"})
        SumInfoList.Add(ComConst.調査区分.経営分析調査_肥育豚生産費, list)

        Return SumInfoList
    End Function

    ''' <summary>
    ''' 個別結果表４項目合計用の定義取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetKangenShiryoSumInfo2() As Dictionary(Of String, List(Of 生産費.sheet2.sumInfo2))
        Dim SumInfoList As New Dictionary(Of String, List(Of 生産費.sheet2.sumInfo2))
        Dim list As New List(Of 生産費.sheet2.sumInfo2)

        list.Clear()

        '肥育豚生産費(個別)
        list = New List(Of 生産費.sheet2.sumInfo2)
        list.Add(New 生産費.sheet2.sumInfo2 With {.ROW = 3, .COL = 3, .KOBETSU_KOBAN1 = "K010847", .KOBETSU_KOBAN2 = "K010848", .KOBETSU_KOBAN3 = "K010849", .KOBETSU_KOBAN4 = "K010850"})
        SumInfoList.Add(ComConst.調査区分.肥育豚生産費統計_個別, list)

        '肥育豚生産費(組織法人)
        list = New List(Of 生産費.sheet2.sumInfo2)
        list.Add(New 生産費.sheet2.sumInfo2 With {.ROW = 3, .COL = 3, .KOBETSU_KOBAN1 = "K010847", .KOBETSU_KOBAN2 = "K010848", .KOBETSU_KOBAN3 = "K010849", .KOBETSU_KOBAN4 = "K010850"})
        SumInfoList.Add(ComConst.調査区分.経営分析調査_肥育豚生産費, list)

        Return SumInfoList
    End Function

    ''' <summary>
    ''' 個別結果表項目の合計算出
    ''' </summary>
    ''' <param name="value1"></param>
    ''' <param name="value2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddKobetsuValue(ParamArray value As String()) As String
        Dim ret As String
        Dim kobetsuValue As Decimal = 0
        Dim work As Decimal

        For Each Val As String In value
            If Decimal.TryParse(Val, work) Then
                kobetsuValue += work
            End If
        Next

        If kobetsuValue = 0 Then
            ret = String.Empty
        Else
            ret = CStr(kobetsuValue)
        End If

        Return ret
    End Function


    ''' <summary>
    ''' 還元資料固定項目の設定(DBで定義している個所以外の設定)
    ''' </summary>
    ''' <param name="sheetNo"></param>
    ''' <param name="arrayNo"></param>
    ''' <param name="kobetsuArray"></param>
    ''' <param name="shukeiArray"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub SetKangenShiryoFixedAdditional(sheetNo As Integer, arrayNo As Integer, ByRef kobetsuArray(,) As Object, ByRef shukeiArray(,) As Object)
        'MyBase.SetKangenShiryoFixedAdditional(sheetNo, arrayNo, kobetsuArray, shukeiArray)
        If sheetNo = 2 Or sheetNo = 3 Then
            If arrayNo = 1 Then
                '還元資料固定項目定義マスタで保持していない項目の設定

                Dim dcSumInfo As Dictionary(Of String, List(Of 生産費.sheet2.SumInfo)) = Me.GetKangenShiryoSumInfo()
                Dim listSumInfo As List(Of 生産費.sheet2.SumInfo) = dcSumInfo(CommonInfo.Chosakubun)

                Dim kobetsuValue As Decimal = 0

                Dim zennen As Boolean = False
                If mChosanenKey.Count <> 1 Then
                    zennen = True
                End If

                For Each sumInfo As 生産費.sheet2.SumInfo In listSumInfo
                    ' 個別 当年 
                    kobetsuArray(sumInfo.ROW, 0) = AddKobetsuValue(
                        mKobetsu(CStr(mChosanenKey(0))).Item(sumInfo.KOBETSU_KOBAN1).値,
                        mKobetsu(CStr(mChosanenKey(0))).Item(sumInfo.KOBETSU_KOBAN2).値)

                    ' 個別 前年
                    If zennen Then
                        kobetsuArray(sumInfo.ROW, 1) = AddKobetsuValue(
                            mKobetsu(CStr(mChosanenKey(1))).Item(sumInfo.KOBETSU_KOBAN1).値,
                            mKobetsu(CStr(mChosanenKey(1))).Item(sumInfo.KOBETSU_KOBAN2).値)
                    End If

                    ' 集計
                    If Not shukeiArray Is Nothing Then
                        '当年
                        shukeiArray(sumInfo.ROW, 0) = AddKobetsuValue(
                            ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN1, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１),
                            ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN2, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１))
                        '前年
                        shukeiArray(sumInfo.ROW, 1) = AddKobetsuValue(
                            ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN1, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年),
                            ShukeiDataHenkan(sumInfo.KOBETSU_KOBAN2, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年))
                    End If

                Next
            End If
        End If

    End Sub

End Class
