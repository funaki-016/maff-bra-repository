Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Data.SqlTypes

''' -----------------------------------------------------------------------------
''' Class	 : DBAccess
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' データベース操作
''' </summary>
''' <remarks>
''' <list type="definition">
'''    <item>
'''       <description>データベース操作を処理する。</description>
'''       <description>また、Disposeメソッドが実行される事でＤＢ接続が解除される。</description>
'''    </item>
''' </list>
''' </remarks>
''' <example>
''' <code >
''' Private Function func()
'''     Dim accs As DBAccess = Nothing
'''
'''     Try
'''
'''         accs = New DBAccess
'''
'''         '結果を戻さない場合の処理(UPDATE,INSERT)
'''         accs.ExecuteNonQuery(sql)
'''
'''         '１列１行目の値のみを取得する。(SELECT COUNT(*) ...)
'''         accs.ExecuteScalar(sql)
'''
'''         'DataReaderで読む(SELECT)
'''         Dim dr As SqlDataReader
'''         dr = accs.ExecuteReader(sql)
'''
'''         'DataSetで取得(SELECT 結果をグリッドに一覧表示をする場合等)
'''         Dim ds As DataSet
'''         ds = accs.GetData(sql, "foo")
'''
'''         ....
'''
'''     Finally
'''         '終了処理
'''
'''         If (Not accs Is Nothing) Then
'''             accs.Dispose()
'''             accs = Nothing
'''         End If
'''     End Try
'''
''' End Function
''' </code >
''' </example>
''' <history>
''' </history>
''' -----------------------------------------------------------------------------
Public Class DBAccess
    Implements IDisposable

#Region "クラス"

    ''' <summary>
    ''' Oracleパラメータ構造体
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Parameter
        ''' <summary>
        ''' パラメータ名
        ''' </summary>
        ''' <remarks></remarks>
        Public name As String

        ''' <summary>
        ''' データ型
        ''' </summary>
        ''' <remarks></remarks>
        Public type As SqlDbType

        ''' <summary>
        ''' 値
        ''' </summary>
        ''' <remarks></remarks>
        Public value As Object

    End Class

#End Region

#Region "メンバ変数"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' DB接続管理クラス
    ''' </summary>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Protected m_Conn As DBConnect

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' SELECT文実行用コマンド
    ''' </summary>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Protected m_Cmd As SqlCommand

#End Region

#Region "プロパティ"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 接続オブジェクト
    ''' </summary>
    ''' <value>接続オブジェクト</value>
    ''' <history>
    ''' 
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Connection() As SqlConnection
        Get
            Return (m_Conn.Connection)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクションオブジェクト
    ''' </summary>
    ''' <value>トランザクションオブジェクト</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Transaction() As SqlTransaction
        Get
            Return (m_Conn.Transaction)
        End Get
    End Property

#End Region

#Region "コンストラクタ、ファイナライザ"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="connect"></param>
    ''' <remarks></remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(connect As String)

        m_Conn = New DBConnect(connect)
        m_Cmd = m_Conn.Connection.CreateCommand()

        m_Cmd.CommandTimeout = DBConnect.cnsConnectTimeout
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="server"></param>
    ''' <param name="db"></param>
    ''' <param name="user"></param>
    ''' <param name="pass"></param>
    ''' <remarks></remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(server As String, db As String, user As String, pass As String)

        m_Conn = New DBConnect(server, db, user, pass)
        m_Cmd = m_Conn.Connection.CreateCommand()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ファイナライザ
    ''' </summary>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub Finalize()

        'Dispose()

        MyBase.Finalize()
    End Sub

#End Region

#Region "公開メソッド"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' リソース解放
    ''' </summary>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Dispose() Implements System.IDisposable.Dispose

        Try
            If (Not m_Cmd Is Nothing) Then
                m_Cmd.Dispose()
            End If

            If (Not m_Conn Is Nothing) Then
                m_Conn.CommitTrans()
                m_Conn.DisConnectDB()
            End If
        Finally
            m_Conn = Nothing
            m_Cmd = Nothing
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' SQL実行
    ''' </summary>
    ''' <param name="SQL">実行するSQL文</param>
    ''' <returns>
    ''' <list type="definition">
    '''    <item>
    '''       <term>SQLの内容により下記の通り</term>
    '''       <description>UPDATE、INSERT、DELETE:  そのコマンドの影響を受ける行の数</description>
    '''       <description>その他:                  -1</description>
    '''    </item>
    ''' </list>
    ''' </returns>
    ''' <remarks>
    ''' SQLを実行実行し、SQLに影響を受ける行数を戻す。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ExecuteNonQuery(ByVal SQL As String, Optional para As List(Of Parameter) = Nothing, Optional logOutput As Boolean = True) As Integer

        Dim cnt As Integer = 0

        Try
            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            If Not para Is Nothing Then
                For i As Integer = 0 To para.Count - 1
                    m_Cmd.Parameters.Add(para.Item(i).name, para.Item(i).type).Value = ConvertValue(para.Item(i).type, para.Item(i).value)
                Next
            End If

            m_Cmd.CommandText = SQL

            If logOutput Then
                'SQL文デバッグ出力
                Trace.WriteLine(String.Format("ExecuteNonQuery:{0}{1}{2}", vbCrLf, ConvertBindToValue(m_Cmd, SQL), vbCrLf))

                'SQL文システムログ出力
                OutputLog.WriteSystemLog(OutputLog.LogLevel.Info, OutputLog.MSGID_SQL, OutputLog.MSG_SQL & vbCrLf & ConvertBindToValue(m_Cmd, SQL))
            End If

            'SQL実行
            cnt = m_Cmd.ExecuteNonQuery

        Finally
            'バインド変数クリア
            m_Cmd.Parameters.Clear()
        End Try

        Return (cnt)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' SQL実行
    ''' </summary>
    ''' <param name="SQL">実行するSQL文</param>
    ''' <returns>オラクル用DataReaderオブジェクト</returns>
    ''' <remarks>
    ''' SQLを実行する。
    ''' 引数SQLで指定されたSQL文を実行し、DataReaderオブジェクトを戻す。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ExecuteReader(ByVal SQL As String, Optional para As List(Of Parameter) = Nothing) As SqlDataReader

        Dim ret As SqlDataReader = Nothing

        Try
            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            If Not para Is Nothing Then
                For i As Integer = 0 To para.Count - 1
                    m_Cmd.Parameters.Add(para.Item(i).name, para.Item(i).type).Value = ConvertValue(para.Item(i).type, para.Item(i).value)
                Next
            End If

            m_Cmd.CommandText = SQL

            'SQL文デバッグ出力
            Trace.WriteLine(String.Format("ExecuteReader:{0}{1}{2}", vbCrLf, ConvertBindToValue(m_Cmd, SQL), vbCrLf))

            'SQL文システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Info, OutputLog.MSGID_SQL, OutputLog.MSG_SQL & vbCrLf & ConvertBindToValue(m_Cmd, SQL))

            'SQL実行
            ret = m_Cmd.ExecuteReader

        Finally
            'バインド変数クリア
            m_Cmd.Parameters.Clear()
        End Try

        Return (ret)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' SQL実行
    ''' </summary>
    ''' <param name="SQL">実行するSQL文</param>
    ''' <returns>結果セットの最初の行の最初の列</returns>
    ''' <remarks>
    ''' <list type="definition">
    '''    <item>
    '''       <term> SQLを実行する。</term>
    '''       <description>結果セットの最初の行の最初の列を返す。</description>
    '''       <description>SELECT COUNT(..), MAX/MIN 系用</description>
    '''    </item>
    ''' </list>
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ExecuteScalar(ByVal SQL As String, Optional para As List(Of Parameter) = Nothing) As Object

        Dim ret As Object

        Try
            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            If Not para Is Nothing Then
                For i As Integer = 0 To para.Count - 1
                    m_Cmd.Parameters.Add(para.Item(i).name, para.Item(i).type).Value = ConvertValue(para.Item(i).type, para.Item(i).value)
                Next
            End If

            m_Cmd.CommandText = SQL

            'SQL文デバッグ出力
            Trace.WriteLine(String.Format("ExecuteScalar:{0}{1}{2}", vbCrLf, ConvertBindToValue(m_Cmd, SQL), vbCrLf))

            'SQL文システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Info, OutputLog.MSGID_SQL, OutputLog.MSG_SQL & vbCrLf & ConvertBindToValue(m_Cmd, SQL))

            'SQL実行
            ret = m_Cmd.ExecuteScalar

        Finally
            'バインド変数クリア
            m_Cmd.Parameters.Clear()
        End Try

        Return (ret)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' SQL実行(DataSet取得)
    ''' </summary>
    ''' <param name="SQL">SQL:SELECT文</param>
    ''' <param name="DsName">データセット名称</param>
    ''' <returns>DataSet SQLSelectにて指定されたSQL実行結果</returns>
    ''' <remarks>
    ''' 引数SQLで指定されたSELECT文を実行し、結果をDataSetに格納して戻す。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetDataSet(ByVal SQL As String, ByVal DsName As String, Optional para As List(Of Parameter) = Nothing) As DataSet

        Dim ret As DataSet = Nothing
        Dim adpt As SqlDataAdapter = Nothing

        Try
            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            If Not para Is Nothing Then
                For i As Integer = 0 To para.Count - 1
                    m_Cmd.Parameters.Add(para.Item(i).name, para.Item(i).type).Value = ConvertValue(para.Item(i).type, para.Item(i).value)
                Next
            End If

            m_Cmd.CommandText = SQL
            adpt = New SqlDataAdapter(m_Cmd)

            'SQL文デバッグ出力
            Trace.WriteLine(String.Format("GetDataSet:{0}{1}{2}", vbCrLf, ConvertBindToValue(m_Cmd, SQL), vbCrLf))

            'SQL文システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Info, OutputLog.MSGID_SQL, OutputLog.MSG_SQL & vbCrLf & ConvertBindToValue(m_Cmd, SQL))

            ret = New DataSet(DsName)
            adpt.Fill(ret)

        Finally
            'バインド変数クリア
            m_Cmd.Parameters.Clear()
            If (Not adpt Is Nothing) Then
                adpt.Dispose()
                adpt = Nothing
            End If
        End Try

        Return (ret)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' SQL実行(DataTable取得)
    ''' </summary>
    ''' <param name="SQL">SQL:SELECT文</param>
    ''' <returns>DataTable SQLSelectにて指定されたSQL実行結果</returns>
    ''' <remarks>
    ''' 引数SQLで指定されたSELECT文を実行し、結果をDataTableに格納して戻す。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetDataTable(ByVal SQL As String, Optional para As List(Of Parameter) = Nothing) As DataTable

        Dim ret As DataTable = Nothing
        Dim adpt As SqlDataAdapter = Nothing

        Try
            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            If Not para Is Nothing Then
                For i As Integer = 0 To para.Count - 1
                    m_Cmd.Parameters.Add(para.Item(i).name, para.Item(i).type).Value = ConvertValue(para.Item(i).type, para.Item(i).value)
                Next
            End If

            m_Cmd.CommandText = SQL
            adpt = New SqlDataAdapter(m_Cmd)

            'SQL文デバッグ出力
            Trace.WriteLine(String.Format("GetDataTable:{0}{1}{2}", vbCrLf, ConvertBindToValue(m_Cmd, SQL), vbCrLf))

            'SQL文システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Info, OutputLog.MSGID_SQL, OutputLog.MSG_SQL & vbCrLf & ConvertBindToValue(m_Cmd, SQL))

            ret = New DataTable
            adpt.Fill(ret)

        Finally
            'バインド変数クリア
            m_Cmd.Parameters.Clear()
            If (Not adpt Is Nothing) Then
                adpt.Dispose()
                adpt = Nothing
            End If
        End Try

        Return (ret)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ロックSQL実行
    ''' </summary>
    ''' <param name="SQL">実行するSQL文</param>
    ''' <returns>ロック成功に成功した場合はTrue</returns>
    ''' <remarks>
    ''' ロックSQLを実行する。
    ''' 引数SQLで指定されたロックSQL文を実行し、結果を戻す。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ExecuteLock(ByVal SQL As String, Optional para As List(Of Parameter) = Nothing) As Boolean

        Dim ret As Boolean = False
        Dim reader As SqlDataReader = Nothing

        Try
            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            If Not para Is Nothing Then
                For i As Integer = 0 To para.Count - 1
                    m_Cmd.Parameters.Add(para.Item(i).name, para.Item(i).type).Value = ConvertValue(para.Item(i).type, para.Item(i).value)
                Next
            End If

            m_Cmd.CommandText = SQL

            'SQL文デバッグ出力
            Trace.WriteLine(String.Format("ExecuteLock:{0}{1}{2}", vbCrLf, ConvertBindToValue(m_Cmd, SQL), vbCrLf))

            'SQL文システムログ出力
            OutputLog.WriteSystemLog(OutputLog.LogLevel.Info, OutputLog.MSGID_SQL, OutputLog.MSG_SQL & vbCrLf & ConvertBindToValue(m_Cmd, SQL))

            'SQL実行
            reader = m_Cmd.ExecuteReader
            While reader.Read()
                ' リードしないとロックされているかわからない
            End While

            ret = True

        Catch ex As Exception
            'ロックの取得エラーの場合にFalseを返すためにCatchする
        Finally
            If Not IsNothing(reader) Then
                reader.Close()
            End If
            'バインド変数クリア
            m_Cmd.Parameters.Clear()
        End Try

        Return (ret)

    End Function

    Private Function ConvertValue(type As SqlDbType, value As Object) As Object
        Dim obj As Object = Nothing

        Select Case (type)
            Case SqlDbType.Int
                Dim cnvInteger As Integer
                If Integer.TryParse(CStr(value), cnvInteger) Then
                    obj = CType(cnvInteger, SqlTypes.SqlInt32)
                Else
                    obj = SqlTypes.SqlInt32.Null
                End If

            Case SqlDbType.Decimal
                Dim cnvDecimal As Decimal
                If Decimal.TryParse(CStr(value), cnvDecimal) Then
                    obj = CType(cnvDecimal, SqlTypes.SqlDecimal)
                Else
                    obj = SqlTypes.SqlDecimal.Null
                End If

            Case SqlDbType.DateTime
                Dim cnvDateTime As DateTime

                If DateTime.TryParse(CStr(value), cnvDateTime) Then
                    obj = CType(cnvDateTime, SqlTypes.SqlDateTime)
                Else
                    obj = SqlTypes.SqlDateTime.Null
                End If

            Case SqlDbType.NVarChar, SqlDbType.VarChar
                If value Is Nothing Then
                    obj = SqlTypes.SqlString.Null
                Else
                    obj = String.Format("{0}", value)
                End If

            Case Else
                Debug.Assert(False, "Parameters Error!")
        End Select

        Return obj

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクション開始処理
    ''' </summary>
    ''' <remarks>
    ''' <list type="definition">
    '''    <item>
    '''       <term>トランザクション処理を開始する。</term>
    '''       <description>データベースが閉じられている場合は、自動オープンする。</description>
    '''       <description>既に、トランザクション処理が開始されている場合、
    '''                    既存のトランザクション処理のコミット処理を行った後に、
    '''                    新規トランザクション処理を開始する。
    '''       </description>
    '''    </item>
    ''' </list>
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginTrans()

        If (Not m_Conn Is Nothing) Then

            If (m_Conn.ConnectState = False) Then
                m_Conn.ConnectDB()
            End If

            'トランザクション処理中の場合には、コミットする。
            m_Conn.CommitTrans()

            'トランザクションを開始する。
            m_Conn.BeginTrans()

            m_Cmd.Transaction = m_Conn.Transaction

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクションコミット処理
    ''' </summary>
    ''' <remarks>
    ''' トランザクション処理をコミットする。
    ''' コミット処理が失敗した場合には、ロールバック処理を行う。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub CommitTrans()

        Try
            If (Not m_Conn Is Nothing) Then
                m_Conn.CommitTrans()
            End If
        Catch ex As Exception
            If (Not m_Conn Is Nothing) Then
                m_Conn.RollBackTrans()
            End If

            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクションロールバック処理
    ''' </summary>
    ''' <remarks>
    ''' トランザクションロールバック処理を実行する。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub RollBackTrans()

        If (Not m_Conn Is Nothing) Then
            m_Conn.RollBackTrans()
        End If

    End Sub

    ''' <summary>
    ''' SQLパラメータを生成して返す
    ''' </summary>
    ''' <param name="name">パラメータ名</param>
    ''' <param name="type">データ型</param>
    ''' <param name="value">値</param>
    ''' <returns>SQLパラメータ</returns>
    ''' <remarks></remarks>
    Public Function CreateParameter(ByVal name As String, ByVal type As SqlDbType, ByVal value As Object) As Parameter

        Dim db As Parameter = Nothing

        Try

            db = New Parameter

            db.name = name
            db.type = type
            db.value = value

        Catch ex As Exception
            Throw ex
        End Try

        Return db

    End Function

    ''' <summary>
    ''' バインド変数を値に置換して出力する
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertBindToValue(cmd As SqlCommand, SQL As String) As String

        'ParameterNameをlistに格納
        Dim list As New ArrayList
        For i = 0 To cmd.Parameters.Count - 1
            If Not cmd.Parameters.Item(i).Value Is Nothing Then
                list.Add(cmd.Parameters.Item(i).ParameterName)
            End If
        Next

        Dim strlenComp As IComparer = New StrLenComparer()
        '文字列の長さの長い順にソート
        list.Sort(strlenComp)

        Dim key As String = String.Empty
        'ParameterNameが長いものから置換する
        For i = 0 To list.Count - 1
            key = list.Item(i).ToString
            Select Case (cmd.Parameters.Item(key).DbType)
                Case DbType.Int32, DbType.Decimal
                    SQL = Replace(SQL, cmd.Parameters.Item(key).ParameterName, cmd.Parameters.Item(key).Value.ToString)
                Case Else
                    SQL = Replace(SQL, cmd.Parameters.Item(key).ParameterName, "'" & cmd.Parameters.Item(key).Value.ToString & "'")
            End Select
        Next

        Return SQL

    End Function

#End Region

    Public Class StrLenComparer
        Implements IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim str1 As String = CType(x, String)
        Dim str2 As String = CType(y, String)

        Return str2.Length - str1.Length
        End Function
    End Class

End Class
