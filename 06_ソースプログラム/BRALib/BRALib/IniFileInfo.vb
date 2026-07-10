Imports System.Runtime.InteropServices
Imports System.Text

Public Class IniFileInfo

    ''' <summary>設定ファイル</summary>
    Private Shared _iniFile As String = String.Empty

    ''' <summary>
    ''' 設定ファイル
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property IniFile() As String
        Set(ByVal value As String)
            _iniFile = value
        End Set
        Get
            Return _iniFile
        End Get
    End Property

    ''' <summary>
    ''' 調査票OCRデータ格納フォルダ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CsvOcrInPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "CSV"
        Dim key As String = "OCRINPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 調査票オンラインデータ格納フォルダ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CsvOnlineInPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "CSV"
        Dim key As String = "ONLINEINPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' EXCEL入力画面フォルダ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExcelScreenPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "EXCEL"
        Dim key As String = "SCREEN"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' EXCEL帳票雛形フォルダ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExcelReportPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "EXCEL"
        Dim key As String = "REPORT"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' EXCEL帳票格納フォルダ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExcelOutPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "EXCEL"
        Dim key As String = "OUTPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
                result.Replace("\(kyoten)", String.Empty)
            Else
                result.Replace("(kyoten)", Mid(CommonInfo.UserOffice, 2))
            End If
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' EXCEL帳票用データ格納フォルダ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExcelInPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "EXCEL"
        Dim key As String = "INPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
                result.Replace("\(kyoten)", String.Empty)
            Else
                result.Replace("(kyoten)", Mid(CommonInfo.UserOffice, 2))
            End If
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' BCP(Bulk Copy Program)ユーティリティの実行ファイルパス
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function BcpExePath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "DB"
        Dim key As String = "BCPEXEPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' ユーザログファイルパス
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogUserLogFilePath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOG"
        Dim key As String = "USERLOGPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(date)", Now.ToString("yyyyMMdd"))
            result.Replace("(time)", Now.ToString("HHmmss"))
            result.Replace("(user)", CommonInfo.UserId)
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' システムログファイルパス
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogSystemLogFilePath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOG"
        Dim key As String = "SYSTEMLOGPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(date)", Now.ToString("yyyyMMdd"))
            result.Replace("(time)", Now.ToString("HHmmss"))
            result.Replace("(kotei)", CommonInfo.Koutei)
            result.Replace("(kyoten)", Mid(CommonInfo.UserOffice, 2))
            result.Replace("(user)", CommonInfo.UserId)
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' ユーザログ出力レベル
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogUserLogLevel() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOG"
        Dim key As String = "USERLOGLEVEL"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' システムログ出力レベル
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogSystemLogLevel() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOG"
        Dim key As String = "SYSTEMLOGLEVEL"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 出力制限サイズ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogMaxSize() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOG"
        Dim key As String = "MAXSIZE"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 履歴件数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogHistoryCount() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOG"
        Dim key As String = "HISTORYCOUNT"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 専門調査員EXCEL帳票用データ格納フォルダ
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SenmonchosainInPath(Optional user As String = Nothing) As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "SENMONCHOSAIN"
        Dim key As String = "INPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            If user Is Nothing Then
                result.Replace("(user)", CommonInfo.UserId)
            Else
                result.Replace("(user)", user)
            End If
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 専門調査員EXCEL帳票格納フォルダ
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SenmonchosainOutPath(Optional user As String = Nothing) As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "SENMONCHOSAIN"
        Dim key As String = "OUTPATH"

        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            If user Is Nothing Then
                result.Replace("(user)", CommonInfo.UserId)
            Else
                result.Replace("(user)", user)
            End If
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' バックアップフォルダ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function BackUpPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "BACKUP"
        Dim key As String = "PATH"

        If GetPrivateProfileString(section, key, String.Empty,
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
                result.Replace("\(kyoten)", String.Empty)
            Else
                result.Replace("(kyoten)", Mid(CommonInfo.UserOffice, 2))
            End If
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' ロックフォルダの上位フォルダ取得
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function LockParentPath() As String
        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOCK"
        Dim key As String = "LOCKPARENTPATH"

        If GetPrivateProfileString(section, key, String.Empty,
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            Return result.ToString
        Else
            Return String.Empty
        End If
    End Function

    ''' <summary>
    ''' ロックフォルダ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LockPath() As String

        Dim result As StringBuilder = New StringBuilder(1024)

        Dim section As String = "LOCK"
        Dim key As String = "LOCKFILEPATH"

        If GetPrivateProfileString(section, key, String.Empty,
                        result, result.Capacity, _iniFile) > 0 Then
            '文字列置換
            result.Replace("(kotei)", CommonInfo.Koutei)
            If CommonInfo.Koutei = CommonInfo.KouteiKubun.Code.Honsyo Then
                result.Replace("\(kyoten)", String.Empty)
            Else
                result.Replace("(kyoten)", Mid(CommonInfo.UserOffice, 2))
            End If
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 設定ファイルキー関連文字列取得
    ''' </summary>
    ''' <param name="section">セクション</param>
    ''' <param name="key">キー</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetProfileString(ByVal section As String, ByVal key As String) As String

        Dim result As StringBuilder = New StringBuilder(1024)

        'API関数呼び出し
        If GetPrivateProfileString(section, key, String.Empty, _
                        result, result.Capacity, _iniFile) > 0 Then
            Return result.ToString
        Else
            Return String.Empty
        End If

    End Function

    ''' <summary>
    ''' 設定ファイルキー関連文字列取得
    ''' </summary>
    ''' <param name="lpAppName">セクション名</param>
    ''' <param name="lpKeyName">キー名</param>
    ''' <param name="lpDefault">既定の文字列</param>
    ''' <param name="lpReturnedString">情報が格納されるバッファ</param>
    ''' <param name="nSize">情報バッファのサイズ</param>
    ''' <param name="lpFileName">.ini ファイルの名前</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)> _
    Private Shared Function GetPrivateProfileString( _
            ByVal lpAppName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As StringBuilder, _
            ByVal nSize As Integer, _
            ByVal lpFileName As String) As Integer
    End Function
End Class
