'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者             | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2019/05/15 | Daiko              | 新規作成
'------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

''' <summary>
''' マスタテーブル操作
''' </summary>
''' <remarks></remarks>
Public Class MasterDao

    ''' <summary>
    ''' マスタ取得
    ''' </summary>
    ''' <param name="masterName">マスタ名称</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetMaster(ByVal masterName As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            Using db As New DBAccess(My.Settings.BRAHConnectionString)

                sb = New System.Text.StringBuilder

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT * ")
                    .AppendLine(String.Format("FROM   ""{0}""", masterName))
                End With

                ret = db.GetDataTable(sb.ToString)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 地域名マスタ取得(北海道、沖縄追加)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetChikiMaster() As DataTable
        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)
                sb = New System.Text.StringBuilder
                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   地域名マスタ")
                    .AppendLine("WHERE  農業地域番号 >= 60 OR 農業地域番号 IN (1,47) ")
                    .AppendLine("ORDER BY 農業地域番号")
                End With
                ret = db.GetDataTable(sb.ToString)
            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 局名を取得する
    ''' </summary>
    ''' <param name="kyokuCode">局コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKyokuName(ByVal kyokuCode As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            If kyokuCode = String.Empty Then
                Return strReturn
            End If

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT 局名")
                    .AppendLine("FROM   局名マスタ")
                    .AppendLine("WHERE  局コード = @局コード")
                End With

                para.Add(db.CreateParameter("@局コード", SqlDbType.Int, kyokuCode))

                dr = db.ExecuteReader(sb.ToString, para)
                If dr.Read Then
                    strReturn = dr("局名").ToString
                End If

            End Using

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn

    End Function

    ''' <summary>
    ''' 事務所名を取得する
    ''' </summary>
    ''' <param name="jimusyoNo">事務所番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetJimusyoName(ByVal jimusyoNo As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            If jimusyoNo = String.Empty Then
                Return strReturn
            End If

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT 事務所名")
                    .AppendLine("FROM   事務所名マスタ")
                    .AppendLine("WHERE  事務所番号 = @事務所番号")
                End With

                para.Add(db.CreateParameter("@事務所番号", SqlDbType.Int, jimusyoNo))

                dr = db.ExecuteReader(sb.ToString, para)
                If dr.Read Then
                    strReturn = dr("事務所名").ToString
                End If

            End Using

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn

    End Function

    ''' <summary>
    ''' センター名を取得する
    ''' </summary>
    ''' <param name="jimusyoNo">事務所番号</param>
    ''' <param name="jimusyoNo">センター番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCenterName(ByVal jimusyoNo As String, ByVal centerNo As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            If jimusyoNo = String.Empty Then
                Return strReturn
            End If

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT センター名")
                    .AppendLine("FROM   センター名マスタ")
                    .AppendLine("WHERE  事務所番号 = @事務所番号")
                    .AppendLine("AND    センター番号 = @センター番号")
                End With

                para.Add(db.CreateParameter("@事務所番号", SqlDbType.Int, jimusyoNo))
                para.Add(db.CreateParameter("@センター番号", SqlDbType.Int, centerNo))

                dr = db.ExecuteReader(sb.ToString, para)
                If dr.Read Then
                    strReturn = dr("センター名").ToString
                End If

            End Using

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn

    End Function

    ''' <summary>
    ''' 局マスタ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKyokuMaster() As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   局名マスタ")
                End With

                ret = db.GetDataTable(sb.ToString)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 事務所マスタ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetJimusyoMaster() As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   事務所名マスタ")
                    .AppendLine("WHERE  事務所番号 <= 51")
                    .AppendLine("ORDER BY CASE 事務所番号 WHEN 51 THEN 1 ELSE 事務所番号 END")
                End With

                ret = db.GetDataTable(sb.ToString, para)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 事務所マスタ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetJimusyoMaster(ByVal kyokuNo As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   事務所名マスタ")
                    .AppendLine("WHERE  局番号 = @局番号")
                    .AppendLine("AND    事務所番号 <= 51")
                End With

                para.Add(db.CreateParameter("@局番号", SqlDbType.Int, kyokuNo))

                ret = db.GetDataTable(sb.ToString, para)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function


    ''' <summary>
    ''' 事務所マスタ取得(東海農政局の場合静岡も取得する。)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetJimusyoMasterShukei(ByVal kyokuNo As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   事務所名マスタ")
                    If kyokuNo = "5" Then
                        .AppendLine("WHERE  (局番号 = @局番号")
                        .AppendLine("OR     事務所番号 = 22)")
                    Else
                        .AppendLine("WHERE  局番号 = @局番号")
                    End If
                    .AppendLine("AND    事務所番号 <= 51")
                End With

                para.Add(db.CreateParameter("@局番号", SqlDbType.Int, kyokuNo))

                ret = db.GetDataTable(sb.ToString, para)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 事務所マスタ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetJimusyoMaster(ByVal kyokuNo As String, ByVal jimusyoNo As String) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   事務所名マスタ")
                    .AppendLine("WHERE  局番号 = @局番号")
                    .AppendLine("AND    事務所番号 = @事務所番号")
                End With

                para.Add(db.CreateParameter("@局番号", SqlDbType.Int, kyokuNo))
                para.Add(db.CreateParameter("@事務所番号", SqlDbType.Int, jimusyoNo))

                ret = db.GetDataTable(sb.ToString, para)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' センター名マスタ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCenterMaster() As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                sb = New System.Text.StringBuilder

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   センター名マスタ")
                End With

                ret = db.GetDataTable(sb.ToString)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' センター名マスタ取得
    ''' </summary>
    ''' <param name="kyoku"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCenterMaster(kyokuNo As String, Optional jimusho As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            Using db As New DBAccess(ComUtil.GetCurrentConnectionString(CommonInfo.Koutei))

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT センター名マスタ.*")
                    .AppendLine("FROM   センター名マスタ")
                    .AppendLine("INNER JOIN COMM.dbo.事務所名マスタ ON センター名マスタ.事務所番号 = 事務所名マスタ.事務所番号")
                    .AppendLine("WHERE  局番号 = @局番号")
                    If Not jimusho Is Nothing Then
                        .AppendLine("AND    センター名マスタ.事務所番号 = @事務所番号 ")
                    End If
                End With

                para.Add(db.CreateParameter("@局番号", SqlDbType.Int, kyokuNo))
                If Not jimusho Is Nothing Then
                    para.Add(db.CreateParameter("@事務所番号", SqlDbType.Int, jimusho))
                End If

                ret = db.GetDataTable(sb.ToString, para)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 都道府名県マスタ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTodofukenMaster() As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT *")
                    .AppendLine("FROM   都道府県名マスタ")
                End With

                ret = db.GetDataTable(sb.ToString)

            End Using

        Catch ex As Exception
            Throw ex
        End Try

        Return ret
    End Function

    ''' <summary>
    ''' 市町村名取得
    ''' </summary>
    ''' <returns>市町村名</returns>
    ''' <remarks></remarks>
    Public Shared Function GetShichosonMaster(Optional ByVal jimushoNo As String = Nothing) As DataTable

        Dim ret As DataTable
        Dim sb As System.Text.StringBuilder = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT   CASE [事務所番号] WHEN 1 THEN 51 ELSE [事務所番号] END AS 事務所番号")
                    .AppendLine("        ,[市町村番号] ")
                    .AppendLine("        ,[市町村名] ")
                    .AppendLine("FROM     [市町村名マスタ] ")
                    .AppendLine("WHERE    [市町村番号] <> 0 ")
                    .AppendLine("AND      [市町村番号] <> 999 ")
                    If Not String.IsNullOrEmpty(jimushoNo) Then
                        .AppendLine("AND      [事務所番号] = @事務所番号 ")
                    End If
                    .AppendLine("ORDER BY CASE [事務所番号] WHEN 51 THEN 1 ELSE [事務所番号] END, [市町村番号] ")
                End With

                If Not String.IsNullOrEmpty(jimushoNo) Then
                    para.Add(db.CreateParameter("@事務所番号", SqlDbType.Int, If(jimushoNo = "51", "1", jimushoNo)))
                End If

                ret = db.GetDataTable(sb.ToString, para)

            End Using

        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

    ''' <summary>
    ''' 局番号を取得する
    ''' </summary>
    ''' <param name="jimusyoNo">事務所番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetKyokuNo(ByVal jimusyoNo As String) As String

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try

            If jimusyoNo = String.Empty Then
                Return strReturn
            End If

            Using db As New DBAccess(My.Settings.COMMConnectionString)

                sb = New System.Text.StringBuilder
                para = New List(Of DBAccess.Parameter)

                ' SQL文の設定
                With sb
                    .AppendLine("SELECT 局番号")
                    .AppendLine("FROM   事務所名マスタ")
                    .AppendLine("WHERE  事務所番号 = @事務所番号")
                End With

                para.Add(db.CreateParameter("@事務所番号", SqlDbType.Int, jimusyoNo))

                dr = db.ExecuteReader(sb.ToString, para)
                If dr.Read Then
                    strReturn = dr("局番号").ToString
                End If

            End Using

        Catch ex As Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Dispose()
            End If
        End Try

        Return strReturn

    End Function

End Class
