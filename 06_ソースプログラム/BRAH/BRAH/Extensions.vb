'//*************************************************************************************************
'//  修正履歴
'// -----------------------------------------------------------------------------------------------
'//  Revision  |   日  付   |       名  前       |                修  正  内  容
'// -----------+------------+--------------------+-------------------------------------------------
'//  REV_000   | 2023.10.30 |大興電子通信        | 要件No.1 新規作成
'//            |            |                    |
'//*************************************************************************************************
Imports System.Runtime.CompilerServices

''' <summary>
''' 拡張メソッドモジュール
''' </summary>
Namespace Extensions

    ''' <summary>
    ''' DataGridViewCell拡張
    ''' </summary>
    Module CellEx
        ''' <summary>
        ''' 値を文字列で取得
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        <Extension>
        Function GetString(cell As DataGridViewCell) As String
            Return If(cell.Value, "").ToString
        End Function
        ''' <summary>
        ''' 値を整数で取得
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        <Extension>
        Function GetInteger(cell As DataGridViewCell) As Integer
            Dim i As Integer
            If Integer.TryParse(cell.GetString, i) Then
                Return i
            End If
            Return 0
        End Function
        ''' <summary>
        ''' 値を小数で取得
        ''' </summary>
        ''' <param name="cell"></param>
        ''' <returns></returns>
        <Extension>
        Function GetDecimal(cell As DataGridViewCell) As Decimal
            Dim d As Decimal
            If Decimal.TryParse(cell.GetString, d) Then
                Return d
            End If
            Return 0D
        End Function
    End Module

    ''' <summary>
    ''' DataRow拡張
    ''' </summary>
    Public Module DataRowEx
        ''' <summary>
        ''' 値を文字列で取得
        ''' </summary>
        ''' <param name="dataRow"></param>
        ''' <param name="columnName"></param>
        ''' <returns></returns>
        <Extension>
        Public Function GetString(dataRow As DataRow, columnName As String) As String
            Dim item = dataRow(columnName)
            If item Is DBNull.Value OrElse item Is Nothing Then
                Return ""
            End If
            Return item.ToString
        End Function
        ''' <summary>
        ''' 値を整数で取得
        ''' </summary>
        ''' <param name="dataRow"></param>
        ''' <param name="columnName"></param>
        ''' <returns></returns>
        <Extension>
        Public Function GetInteger(dataRow As DataRow, columnName As String) As Integer
            Dim i As Integer
            If Integer.TryParse(dataRow.GetString(columnName), i) Then
                Return i
            End If
            Return 0
        End Function
        ''' <summary>
        ''' 値を小数で取得
        ''' </summary>
        ''' <param name="dataRow"></param>
        ''' <param name="columnName"></param>
        ''' <returns></returns>
        <Extension>
        Public Function GetDecimal(dataRow As DataRow, columnName As String) As Decimal
            Dim d As Decimal
            If Decimal.TryParse(dataRow.GetString(columnName), d) Then
                Return d
            End If
            Return 0D
        End Function

    End Module

End Namespace
