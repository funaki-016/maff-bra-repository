Imports System.Data.SqlClient

Public Class CommonInfo

    ''' <summary>
    ''' 工程区分
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KouteiKubun
        ''' <summary>
        ''' 工程区分コード
        ''' </summary>
        ''' <remarks></remarks>
        Public Class Code
            Public Const Honsyo As String = "H"
            Public Const Kyoku As String = "N"
            Public Const Center As String = "S"
        End Class

        ''' <summary>
        ''' 工程区分名
        ''' </summary>
        ''' <remarks></remarks>
        Public Class Name
            Public Const Honsyo As String = "本省工程"
            Public Const Kyoku As String = "農政局工程"
            Public Const Center As String = "実査設置拠点工程"
        End Class
    End Class

    ''' <summary>調査ID</summary>
    Public Const ChosaID As String = "BRA"

    ''' <summary>本省</summary>
    Public Const HONSYONAME As String = "本省"

    ''' <summary>ユーザID</summary>
    Private Shared _UserId As String = Nothing
    ''' <summary>局CD+事務所CD+センターID</summary>
    Private Shared _UserOffice As String = Nothing
    ''' <summary>局</summary>
    Private Shared _Kyoku As String = Nothing
    ''' <summary>局名</summary>
    Private Shared _KyokuName As String = Nothing
    ''' <summary>事務所</summary>
    Private Shared _Jimusyo As String = Nothing
    ''' <summary>事務所名</summary>
    Private Shared _JimusyoName As String = Nothing
    ''' <summary>センター</summary>
    Private Shared _Center As String = Nothing
    ''' <summary>センター名</summary>
    Private Shared _CenterName As String = Nothing
    ''' <summary>工程</summary>
    Private Shared _Koutei As String = Nothing
    ''' <summary>工程名</summary>
    Private Shared _KouteiName As String = Nothing

    ''' <summary>接続文字列</summary>
    Private Shared _Connect As String

    ''' <summary>専門調査員</summary>
    Private Shared _SenmonChosain As Boolean = Nothing

    ''' <summary>区分１</summary>
    Private Shared _Kubun1 As String = Nothing
    ''' <summary>区分１名</summary>
    Private Shared _Kubun1Name As String = Nothing
    ''' <summary>区分２</summary>
    Private Shared _Kubun2 As String = Nothing
    ''' <summary>区分２名</summary>
    Private Shared _Kubun2Name As String = Nothing
    ''' <summary>調査区分</summary>
    Private Shared _Chosakubun As String = Nothing
    ''' <summary>調査区分名</summary>
    Private Shared _ChosakubunName As String = Nothing

    ''' <summary>
    ''' ユーザID
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property UserId() As String
        Set(ByVal value As String)
            _UserId = value
        End Set
        Get
            Return _UserId
        End Get
    End Property

    ''' <summary>
    ''' 局CD+事務所CD+センターID
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property UserOffice() As String
        Set(ByVal value As String)
            _UserOffice = GetUserOffice(_Connect, value)

            _Kyoku = Mid(_UserOffice, 2, 2)
            _Jimusyo = Mid(_UserOffice, 4, 2)
            _Center = Mid(_UserOffice, 6, 2)

            _Koutei = Left(_UserOffice, 1)

            Select Case _Koutei
                Case KouteiKubun.Code.Center
                    _KouteiName = KouteiKubun.Name.Center
                Case KouteiKubun.Code.Kyoku
                    _KouteiName = KouteiKubun.Name.Kyoku
                Case KouteiKubun.Code.Honsyo
                    _KouteiName = KouteiKubun.Name.Honsyo
            End Select
        End Set
        Get
            Return _UserOffice
        End Get
    End Property

    ''' <summary>
    ''' 局
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Kyoku() As String
        Get
            Return _Kyoku
        End Get
    End Property

    ''' <summary>
    ''' 局名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property KyokuName() As String
        Set(ByVal value As String)
            _KyokuName = value
        End Set
        Get
            Return _KyokuName
        End Get
    End Property

    ''' <summary>
    ''' 事務所
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Jimusyo() As String
        Get
            Return _Jimusyo
        End Get
    End Property

    ''' <summary>
    ''' 事務所名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property JimusyoName() As String
        Set(ByVal value As String)
            _JimusyoName = value
        End Set
        Get
            Return _JimusyoName
        End Get
    End Property

    ''' <summary>
    ''' センター
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Center() As String
        Get
            Return _Center
        End Get
    End Property

    ''' <summary>
    ''' センター名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property CenterName() As String
        Set(ByVal value As String)
            _CenterName = value
        End Set
        Get
            Return _CenterName
        End Get
    End Property

    ''' <summary>
    ''' 工程
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Koutei() As String
        Get
            Return _Koutei
        End Get
    End Property

    ''' <summary>
    ''' 工程名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property KouteiName() As String
        Get
            Return _KouteiName
        End Get
    End Property

    ''' <summary>
    ''' 接続文字列
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared WriteOnly Property Connect() As String
        Set(ByVal value As String)
            _Connect = value
        End Set
    End Property

    ''' <summary>
    ''' 専門調査員
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property SenmonChosain() As Boolean
        Set(ByVal value As Boolean)
            _SenmonChosain = value
        End Set
        Get
            Return _SenmonChosain
        End Get
    End Property

    ''' <summary>
    ''' 区分１
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property Kubun1() As String
        Set(ByVal value As String)
            _Kubun1 = value
        End Set
        Get
            Return _Kubun1
        End Get
    End Property

    ''' <summary>
    ''' 区分１名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property Kubun1Name() As String
        Set(ByVal value As String)
            _Kubun1Name = value
        End Set
        Get
            Return _Kubun1Name
        End Get
    End Property

    ''' <summary>
    ''' 区分２
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property Kubun2() As String
        Set(ByVal value As String)
            _Kubun2 = value
        End Set
        Get
            Return _Kubun2
        End Get
    End Property

    ''' <summary>
    ''' 区分２名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property Kubun2Name() As String
        Set(ByVal value As String)
            _Kubun2Name = value
        End Set
        Get
            Return _Kubun2Name
        End Get
    End Property

    ''' <summary>
    ''' 調査区分
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property Chosakubun() As String
        Set(ByVal value As String)
            _Chosakubun = value
        End Set
        Get
            Return _Chosakubun
        End Get
    End Property

    ''' <summary>
    ''' 調査区分名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Property ChosakubunName() As String
        Set(ByVal value As String)
            _ChosakubunName = value
        End Set
        Get
            Return _ChosakubunName
        End Get
    End Property

    ''' <summary>
    ''' 局CD+事務所CD+センターID取得
    ''' </summary>
    ''' <param name="connect"></param>
    ''' <param name="userOffice"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetUserOffice(connect As String, ByVal userOffice As String) As String

        Dim jimusyoNo As String = String.Empty
        Dim centerNO As String = String.Empty

        Dim strReturn As String = String.Empty
        Dim sb As System.Text.StringBuilder = Nothing
        Dim dr As SqlDataReader = Nothing
        Dim para As List(Of DBAccess.Parameter) = Nothing

        Try
            Select Case Left(userOffice, 1)
                Case KouteiKubun.Code.Center

                    jimusyoNo = CStr(CInt(Mid(userOffice, 4, 2)))
                    centerNO = CStr(CInt(Mid(userOffice, 6, 2)))

                    Using db As New DBAccess(connect)

                        sb = New System.Text.StringBuilder
                        para = New List(Of DBAccess.Parameter)

                        ' SQL文の設定
                        With sb
                            .AppendLine("SELECT JIMUSHO")
                            .AppendLine("      ,CENTER")
                            .AppendLine("FROM   北海道コード変換マスタ")
                            .AppendLine("WHERE  COMM_JIMUSHO = @COMM_JIMUSHO")
                            .AppendLine("AND    COMM_CENTER = @COMM_CENTER")
                        End With

                        para.Add(db.CreateParameter("@COMM_JIMUSHO", SqlDbType.Int, jimusyoNo))
                        para.Add(db.CreateParameter("@COMM_CENTER", SqlDbType.Int, centerNO))

                        dr = db.ExecuteReader(sb.ToString, para)
                        If dr.Read Then
                            strReturn = Left(userOffice, 1) & Mid(userOffice, 2, 2) & CInt(dr("JIMUSHO")).ToString("00") & CInt(dr("CENTER")).ToString("00")
                        Else
                            strReturn = userOffice
                        End If

                    End Using
                Case KouteiKubun.Code.Kyoku
                    strReturn = Left(userOffice, 1) & Mid(userOffice, 2, 2) & "0000"
                Case KouteiKubun.Code.Honsyo
                    strReturn = Left(userOffice, 1) & "000000"
            End Select

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
