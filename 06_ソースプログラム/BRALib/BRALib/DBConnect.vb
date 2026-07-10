Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Data.SqlTypes

''' -----------------------------------------------------------------------------
''' Class	 : DBConnect
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' データベース接続管理
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' </history>
''' -----------------------------------------------------------------------------
Public Class DBConnect
    Implements IDisposable

#Region "静的変数"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' コマンドが実行されるまでの待機時間(秒)
    ''' </summary>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Const cnsConnectTimeout As Integer = 3600

#End Region

#Region "メンバ変数"

    Private m_SqlConn As SqlConnection      '接続オブジェクト
    Private m_Trans As SqlTransaction       'トランザクションオブジェクト
    Private m_ServerName As String          'サーバー名
    Private m_DBName As String              'データベース名
    Private m_User As String                'DB接続ユーザー
    Private m_Passwd As String              'DB接続認証パスワード

#End Region

#Region "プロパティ"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' データベース接続状態
    ''' </summary>
    ''' <value>
    ''' <list type="definition">
    '''    <item>
    '''       <term>接続状態</term>
    '''       <description>true:    接続中</description>
    '''       <description>false:   切断中</description>
    '''    </item>
    ''' </list>
    ''' </value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ConnectState() As Boolean
        Get
            Dim flg As Boolean = False

            If (Not m_SqlConn Is Nothing) Then
                If (m_SqlConn.State = ConnectionState.Open) Then
                    flg = True
                End If
            End If

            Return (flg)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 接続オブジェクト
    ''' </summary>
    ''' <value>接続オブジェクト</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Connection() As SqlConnection
        Get
            Return (m_SqlConn)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクション
    ''' </summary>
    ''' <value>トランザクション</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Transaction() As SqlTransaction
        Get
            Return (m_Trans)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' サーバー名
    ''' </summary>
    ''' <value>サーバー名</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ServerName() As String
        Get
            Return (m_ServerName)
        End Get
        Set(ByVal Value As String)
            m_ServerName = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' データベース名
    ''' </summary>
    ''' <value>データベース名</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property DBName() As String
        Get
            Return (m_DBName)
        End Get
        Set(ByVal Value As String)
            m_DBName = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 接続ユーザー名
    ''' </summary>
    ''' <value>接続ユーザー名</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property User() As String
        Get
            Return (m_User)
        End Get
        Set(ByVal Value As String)
            m_User = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 接続ユーザーパスワード
    ''' </summary>
    ''' <value>接続ユーザーパスワード</value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Passwd() As String
        Get
            Return (m_Passwd)
        End Get
        Set(ByVal Value As String)
            m_Passwd = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 接続文字列
    ''' </summary>
    ''' <value>
    ''' <list type="definition">
    '''    <item>
    '''       <term>下記のルールで接続文字列を戻す。</term>
    '''       <description>DB接続中:    接続文字列</description>
    '''       <description>DBクローズ:  空文字</description>
    '''    </item>
    ''' </list>
    ''' </value>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ConnectString() As String

        Get
            Dim ret = String.Empty

            If (Not m_SqlConn Is Nothing) Then
                ret = m_SqlConn.ConnectionString
            End If

            Return (ret)
        End Get
        Set(ByVal Value As String)
            m_SqlConn.ConnectionString = Value
        End Set

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
        m_Trans = Nothing
        m_SqlConn = New SqlConnection
        m_SqlConn.ConnectionString = connect
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
        m_ServerName = server
        m_DBName = db
        m_User = user
        m_Passwd = pass
        m_Trans = Nothing
        m_SqlConn = New SqlConnection
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
    ''' リソースを解放する。
    ''' </summary>
    ''' <remarks>
    ''' データベース接続をクローズし、全リソースを解放する。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Dispose() Implements System.IDisposable.Dispose

        Try
            If (Not m_SqlConn Is Nothing) Then
                If ((m_SqlConn.State = ConnectionState.Open) Or _
                    (m_SqlConn.State = ConnectionState.Broken)) Then

                    m_SqlConn.Close()
                    m_SqlConn.Dispose()
                End If
            End If
        Finally
            m_SqlConn = Nothing
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' データベースに接続する。
    ''' </summary>
    ''' <returns>オラクル接続オブジェクト</returns>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overridable Overloads Function ConnectDB() As SqlConnection

        OpenDB()

        Return (m_SqlConn)
    End Function

    'Public Overloads Function ConnectDB(ByVal CnnSetting As DBCnnSetting) As SqlConnection

    '    If (Not CnnSetting Is Nothing) Then
    '        With CnnSetting
    '            m_User = .Uid
    '            m_Passwd = .Pwd
    '            m_DBName = .DBn
    '        End With
    '    End If

    '    OpenDB()

    '    Return (m_SqlConn)

    'End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' データベース接続クローズ処理
    ''' </summary>
    ''' <remarks>
    ''' 内部でDisposeを呼び出しデータベースリソース解放処理を行う。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub DisConnectDB()

        Dispose()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクション処理を実行する。
    ''' </summary>
    ''' <remarks>
    ''' <list type="definition">
    '''    <item>
    '''       <term>トランザクション処理を実行する。。</term>
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

        If (m_SqlConn Is Nothing) Then
            m_SqlConn = New SqlConnection
        End If

        If (ConnectState = False) Then
            ConnectDB()
        End If

        'トランザクション処理中の場合はコミットする。
        If (Not m_Trans Is Nothing) Then
            If (Not m_Trans Is Nothing) Then
                m_Trans.Commit()
            End If

            m_Trans.Dispose()
            m_Trans = Nothing
        End If

        m_Trans = m_SqlConn.BeginTransaction()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' トランザクション処理コミット
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
            If ((Not m_SqlConn Is Nothing) AndAlso _
                (Not m_Trans Is Nothing)) Then
                m_Trans.Commit()
            End If
        Catch ex As Exception
            'ロールバックを実施する。
            If ((Not m_Trans Is Nothing) And (Not m_Trans.Connection Is Nothing)) Then
                m_Trans.Rollback()
            End If

            Throw ex
        Finally
            If (Not m_Trans Is Nothing) Then
                m_Trans.Dispose()
                m_Trans = Nothing
            End If
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ロールバック
    ''' </summary>
    ''' <remarks>
    ''' ロールバックを実行する。
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub RollBackTrans()

        Try
            If (Not m_Trans Is Nothing) Then
                m_Trans.Rollback()
            End If
        Finally
            If (Not m_Trans Is Nothing) Then
                m_Trans.Dispose()
                m_Trans = Nothing
            End If
        End Try

    End Sub

#End Region

#Region "プライベートメソッド"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' データベースに接続する。
    ''' </summary>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub OpenDB()

        Dim errflg As Boolean = False

        If (m_SqlConn Is Nothing) Then
            m_SqlConn = New SqlConnection
        End If

        '接続中の場合は、再接続は行わない。
        If (ConnectState) Then
            Return
        End If

        Try
            m_SqlConn.ConnectionString = ConnectString

            m_SqlConn.Open()
        Finally
            If (errflg) Then
                Dispose()
            End If
        End Try

    End Sub

#End Region

End Class
