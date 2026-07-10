Public Class ProgressDialog
    Inherits System.Windows.Forms.Form

    'ダイアログフォーム
    Private form As ProgressDialog

    'フォームが表示されるまで待機するための待機ハンドル
    Private startEvent As System.Threading.ManualResetEvent

    'フォームが一度表示されたか
    Private showed As Boolean = False

    'フォームをコードで閉じているか
    Private MeClosing As Boolean = False

    'オーナーフォーム
    Private ownerForm As Form

    '別処理をするためのスレッド
    Private thread As System.Threading.Thread

    'フォームのタイトル
    Private _title As String = "進捗状況表示"

    'ProgressBarの最小、最大、現在の値
    Private _minimum As Integer = 0

    Private _maximum As Integer = 0

    Private _value As Integer = 0

    Private _valueChk As Integer = 0

    Private _select As Integer = 0

    Private _displayMode As Integer = 0

    '表示するメッセージ
    Private _message As String = "処理中・・・"

    Private EndFlg As Boolean = False

    ' Windows フォーム デザイナで必要です。
    Private components2 As System.ComponentModel.IContainer

    Private _contentText As String = ""

    'メッセージボックス表示デリゲート
    Private Delegate Function SetShowMsgBoxDelegate1(ByVal messageId As String, _
                                                     ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                                     ByVal defBtn As MessageBoxDefaultButton) As Microsoft.VisualBasic.MsgBoxResult
    Private Delegate Function SetShowMsgBoxDelegate2(ByVal messageId As String, _
                                                     ByVal msgPara() As String, _
                                                     ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                                     ByVal defBtn As MessageBoxDefaultButton) As Microsoft.VisualBasic.MsgBoxResult

    'メッセージフォーム表示デリゲート
    Private Delegate Sub SetShowMsgFormDelegate(ByVal messageId As String, _
                                                ByVal msgPara() As String, _
                                                ByVal label As String)

#Region " メイン処理 "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

    End Sub

    ''' <summary>
    ''' フォームロード
    ''' </summary>
    Private Sub ProgressForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    ''' <summary>
    ''' プログレスバーのタイトル
    ''' </summary>
    Public Property Title() As String

        Get
            Return _title
        End Get
        Set(ByVal Value As String)
            _title = Value
            If Not (form Is Nothing) Then
                Invoke(New MethodInvoker( _
                    AddressOf SetTitle))
            End If
        End Set

    End Property

    ''' <summary>
    ''' プログレスバーの最小値
    ''' </summary>
    Public Property Minimum() As Integer
        Get
            Return _minimum
        End Get
        Set(ByVal Value As Integer)
            _minimum = Value
            If Not (form Is Nothing) Then
                Invoke(New MethodInvoker( _
                    AddressOf SetProgressMinimum))
            End If
        End Set
    End Property

    ''' <summary>
    ''' プログレスバーの最大値
    ''' </summary>
    Public WriteOnly Property Maximum() As Integer
        Set(ByVal Value As Integer)
            _maximum = Value
            If Not (form Is Nothing) Then
                Invoke(New MethodInvoker( _
                    AddressOf SetProgressMaximun))
            End If
        End Set
    End Property

    ''' <summary>
    ''' プログレスバーの値
    ''' </summary>
    Public WriteOnly Property Value() As Integer
        Set(ByVal Value As Integer)
            _value = Value
            _valueChk = _value
            If Not (form Is Nothing) Then
                If _value <= _maximum Then
                    Invoke(New MethodInvoker(AddressOf SetProgressValue))
                    Invoke(New MethodInvoker(AddressOf SetPBLabelValue))
                Else
                    _value = _valueChk
                End If
            End If
        End Set
    End Property


    ''' <summary>
    ''' プログレスバーの値に足しこむ
    ''' </summary>
    Public WriteOnly Property AddValue() As Integer
        Set(ByVal Value As Integer)
            _value = _value + Value
            _valueChk = _value
            If Not (form Is Nothing) Then
                If _value <= _maximum Then
                    Invoke(New MethodInvoker(AddressOf SetProgressValue))
                    Invoke(New MethodInvoker(AddressOf SetPBLabelValue))
                Else
                    _value = _valueChk
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' 処理内容ラベルを変更する
    ''' </summary>
    Public WriteOnly Property SetProcessText() As String
        Set(ByVal Value As String)
            _contentText = Value
            If Not (form Is Nothing) Then
                Invoke(New MethodInvoker(AddressOf SetProcessTargetLabel))
            End If
        End Set
    End Property

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <param name="style"></param>
    ''' <param name="defBtn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ShowMsgBox(ByVal messageId As String, _
                               ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                               Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult
        Dim ret As Microsoft.VisualBasic.MsgBoxResult
        If Not (form Is Nothing) Then
            ret = DirectCast(Invoke(New SetShowMsgBoxDelegate1(AddressOf SetShowMsgBox), messageId, style, defBtn), Microsoft.VisualBasic.MsgBoxResult)
        End If
        Return ret
    End Function

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <param name="msgPara"></param>
    ''' <param name="style"></param>
    ''' <param name="defBtn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ShowMsgBox(ByVal messageId As String, _
                               ByVal msgPara() As String, _
                               ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                               Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult
        Dim ret As Microsoft.VisualBasic.MsgBoxResult
        If Not (form Is Nothing) Then
            ret = DirectCast(Invoke(New SetShowMsgBoxDelegate2(AddressOf SetShowMsgBox), messageId, msgPara, style, defBtn), Microsoft.VisualBasic.MsgBoxResult)
        End If
        Return ret
    End Function

    ''' <summary>
    ''' メッセージフォームを表示する
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <param name="msgPara"></param>
    ''' <param name="label"></param>
    ''' <remarks></remarks>
    Public Sub ShowMsgForm(ByVal messageId As String, _
                           ByVal msgPara() As String, _
                           Optional ByVal label As String = Nothing)
        If Not (form Is Nothing) Then
            Invoke(New SetShowMsgFormDelegate(AddressOf SetShowMsgForm), messageId, msgPara, label)
        End If
    End Sub

    ''' <summary>
    ''' 終了フラグ
    ''' パラメータ｜True　： 終了する
    ''' 　　　　　｜False ： 終了しない
    ''' </summary>
    Public ReadOnly Property EndValue() As Boolean
        Get
            Return EndFlg
        End Get
    End Property

    ''' <summary>
    ''' ダイアログを表示する
    ''' ownerの中央にダイアログが表示される
    ''' このメソッドは一回しか呼び出せません。
    ''' </summary>
    Public Overloads Sub Show(ByVal owner As IWin32Window)

        If showed Then
            Throw New Exception("ダイアログは一度表示されています。")
        End If

        showed = True
        startEvent = New System.Threading.ManualResetEvent(False)

        '引数型チェック
        If TypeOf owner Is Form Then
            ownerForm = DirectCast(owner, Form)
        End If

        'スレッドを作成
        thread = New System.Threading.Thread( _
            New System.Threading.ThreadStart(AddressOf Run))

        'スレッドをバックグラウンド処理として扱う
        thread.IsBackground = True
        Me.thread.SetApartmentState(Threading.ApartmentState.STA)

        'スレッド開始
        thread.Start()

        'フォームが表示されるまで待機する
        startEvent.WaitOne()

    End Sub

    Public Overloads Sub Show()
        Show(Nothing)
    End Sub

    ''' <summary>
    ''' 別スレッドで処理するメソッド
    ''' </summary>
    Private Sub Run()

        'フォームの設定
        form = New ProgressDialog
        Text = _title
        AddHandler Activated, AddressOf form_Activated
        MainProgressBar.Minimum = _minimum
        MainProgressBar.Maximum = _maximum
        MainProgressBar.Value = _value

        'フォームの表示位置をオーナーの中央へ
        If Not (ownerForm Is Nothing) Then
            StartPosition = FormStartPosition.Manual
            Left = ownerForm.Left + (ownerForm.Width - Width) \ 2
            Top = ownerForm.Top + (ownerForm.Height - Height) \ 2
        End If

        'フォームの表示
        Me.ShowDialog()

    End Sub

    ''' <summary>
    ''' インスタンスを開放する
    ''' </summary>
    Public Sub endDispose()
        'プログレスバーを最大値まで表示させるための調整
        For i As Integer = 1 To 10
            Application.DoEvents()
            System.Threading.Thread.Sleep(100)
        Next

        If showed Then
            showed = False
            Invoke(New MethodInvoker(AddressOf Dispose))
        End If
    End Sub

    ''' <summary>
    ''' プログレスバーの現在値を設定
    ''' </summary>
    Private Sub SetPBLabelValue()

        If Not (form Is Nothing) And Not IsDisposed Then
            'MainProgressBar.Value = _maximum
            MainProgressBar.Value = _value
            If _value >= _maximum Then
                ProcessLabel.Text = _value & "/" & _maximum & " 件"
            End If
        End If

    End Sub

    ''' <summary>
    ''' 処理件数を更新
    ''' </summary>
    Private Sub SetProgressValue()

        If Not (form Is Nothing) And Not IsDisposed Then
            ProcessLabel.Text = _value & "/" & _maximum & " 件"
        End If

    End Sub

    ''' <summary>
    ''' 進捗表示用ラベルに表示するメッセージを設定
    ''' </summary>
    Private Sub SetMessage()
        If Not (form Is Nothing) And Not IsDisposed Then
            ProcessLabel.Text = _message
        End If
    End Sub

    ''' <summary>
    ''' プログレスバーの最大値を設定
    ''' </summary>
    Private Sub SetProgressMaximun()
        If Not (form Is Nothing) And Not IsDisposed Then
            MainProgressBar.Maximum = _maximum
        End If
    End Sub

    ''' <summary>
    ''' プログレスバーの最小値を設定
    ''' </summary>
    Private Sub SetProgressMinimum()
        If Not (form Is Nothing) And Not IsDisposed Then
            MainProgressBar.Minimum = _minimum
        End If
    End Sub

    ''' <summary>
    ''' タイトルを設定
    ''' </summary>
    Private Sub SetTitle()
        If Not (form Is Nothing) And Not IsDisposed Then
            Text = _title
        End If
    End Sub

    ''' <summary>
    ''' フォームの活性処理
    ''' </summary>
    Private Sub form_Activated(ByVal sender As Object, ByVal e As EventArgs)
        RemoveHandler Me.Activated, AddressOf form_Activated
        startEvent.Set()
    End Sub

    ''' <summary>
    ''' Alt + F4 を無効化
    ''' </summary>
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Const WM_SYSCOMMAND As Integer = &H112
        Const SC_CLOSE As Long = &HF060

        If m.Msg = WM_SYSCOMMAND And m.WParam.ToInt64() = SC_CLOSE Then
            Return
        End If

        MyBase.WndProc(m)

    End Sub

    ''' <summary>
    ''' ラベルに表示するメッセージを更新
    ''' </summary>
    Private Sub LabelValue()
        If Not (form Is Nothing) And Not IsDisposed Then

            ProcessTargetLabel.Text = "ファイル情報取得中・・・"

        End If
    End Sub

    ''' <summary>
    ''' 処理内容ラベルを変更する
    ''' </summary>
    Private Sub SetProcessTargetLabel()

        If Not (form Is Nothing) And Not IsDisposed Then
            ProcessTargetLabel.Text = _contentText
        End If

    End Sub

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <param name="style"></param>
    ''' <param name="defBtn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetShowMsgBox(ByVal messageId As String, _
                                   ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                   Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult
        Dim ret As Microsoft.VisualBasic.MsgBoxResult = Nothing
        If Not (form Is Nothing) And Not IsDisposed Then
            ret = Message.ShowMsgBox(messageId, style, defBtn)
        End If
        Return ret
    End Function

    ''' <summary>
    ''' メッセージボックスを表示する
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <param name="msgPara"></param>
    ''' <param name="style"></param>
    ''' <param name="defBtn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetShowMsgBox(ByVal messageId As String, _
                                   ByVal msgPara() As String, _
                                   ByVal style As Microsoft.VisualBasic.MsgBoxStyle, _
                                   Optional ByVal defBtn As MessageBoxDefaultButton = MessageBoxDefaultButton.Button1) As Microsoft.VisualBasic.MsgBoxResult
        Dim ret As Microsoft.VisualBasic.MsgBoxResult = Nothing
        If Not (form Is Nothing) And Not IsDisposed Then
            ret = Message.ShowMsgBox(messageId, msgPara, style, defBtn)
        End If
        Return ret
    End Function

    ''' <summary>
    ''' メッセージフォームを表示する
    ''' </summary>
    ''' <param name="messageId"></param>
    ''' <param name="msgPara"></param>
    ''' <param name="label"></param>
    ''' <remarks></remarks>
    Private Sub SetShowMsgForm(ByVal messageId As String, _
                               ByVal msgPara() As String, _
                               ByVal label As String)
        If Not (form Is Nothing) And Not IsDisposed Then
            Message.ShowMsgForm(Me, messageId, msgPara, label)
        End If
    End Sub
#End Region

End Class
