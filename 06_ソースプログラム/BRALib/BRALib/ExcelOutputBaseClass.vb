Imports Microsoft.Office.Interop
Imports System.IO

''' <summary>
''' Excel出力クラス
''' </summary>
''' <remarks></remarks>
Public MustInherit Class ExcelOutputBaseClass
    Inherits ExcelProcess
    Implements IDisposable

    ''' <summary>出力処理実行結果</summary>
    Public Enum enmOutputReturn
        CANCEL
        OK
        ERR_SAVEAS
    End Enum

    ''' <summary>
    ''' Excel保存エラークラス
    ''' </summary>
    Public Class SaveAsException
        Inherits Exception

        Public Sub New()
            MyBase.New("Excel保存エラー")
        End Sub
    End Class

    ''' <summary>フォント名</summary>
    Protected Const FONT_MS_GOTHIC As String = "ＭＳ ゴシック"
    Protected Const FONT_MS_MINCHO As String = "ＭＳ 明朝"

    ''' <summary>ストップウォッチ：画面表示～画面終了までを計測</summary>
    Protected sw As New System.Diagnostics.Stopwatch()
    ''' <summary>処理名</summary>
    Protected processName As String

    ''' <summary>進捗ダイアログ</summary>
    Protected ProgressDialog As ProgressDialog

    ''' <summary>出力先ディレクトリ</summary>
    Private _OutDir As String
    ''' <summary>出力先パス</summary>
    Private _OutPath As String
    ''' <summary>Excel出力</summary>
    Private _ExcelCheck As Boolean
    ''' <summary>用紙印刷</summary>
    Private _ReportCheck As Boolean
    ''' <summary>マクロ除去</summary>
    Private _RemovalMacro As Boolean

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub New()

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="excelCheck"></param>
    ''' <param name="reportCheck"></param>
    ''' <param name="title"></param>
    ''' <param name="outDir"></param>
    ''' <param name="removalMacro"></param>
    ''' <remarks></remarks>
    Protected Sub New(excelCheck As Boolean, reportCheck As Boolean, ByVal title As String, ByVal outDir As String, ByVal removalMacro As Boolean)

        processName = title
        _ExcelCheck = excelCheck
        _ReportCheck = reportCheck
        _OutDir = outDir
        _RemovalMacro = removalMacro

    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="excelCheck"></param>
    ''' <param name="reportCheck"></param>
    ''' <param name="title"></param>
    ''' <param name="outDir"></param>
    ''' <param name="outPath"></param>
    ''' <param name="removalMacro"></param>
    ''' <remarks></remarks>
    Protected Sub New(excelCheck As Boolean, reportCheck As Boolean, ByVal title As String, ByVal outDir As String, ByVal outPath As String, ByVal removalMacro As Boolean)

        processName = title
        _ExcelCheck = excelCheck
        _ReportCheck = reportCheck
        _OutDir = outDir
        _OutPath = outPath
        _RemovalMacro = removalMacro

    End Sub

    ''' <summary>
    ''' 出力
    ''' </summary>
    ''' <remarks></remarks>
    Public Function OutPut(xlBook As Excel.Workbook, overWriteMessageID As String) As enmOutputReturn
        Dim ret As enmOutputReturn = enmOutputReturn.CANCEL
        If _ExcelCheck Then
            If Not System.IO.Directory.Exists(_OutDir) Then
                Directory.CreateDirectory(_OutDir)
            End If
            If IO.File.Exists(_OutPath) Then
                If Not overWriteMessageID Is Nothing Then
                    Dim msgRet As MsgBoxResult
                    If Not ProgressDialog Is Nothing Then
                        msgRet = ProgressDialog.ShowMsgBox(overWriteMessageID, {IO.Path.GetFileName(_OutPath)}, MsgBoxStyle.YesNo)
                    Else
                        msgRet = Message.ShowMsgBox(overWriteMessageID, {IO.Path.GetFileName(_OutPath)}, MsgBoxStyle.YesNo)
                    End If
                    If msgRet = MsgBoxResult.No Then
                        Return ret
                    End If
                End If
                Try
                    If _RemovalMacro Then
                        xlBook.SaveAs(_OutPath, Excel.XlFileFormat.xlOpenXMLWorkbook)
                    Else
                        xlBook.SaveAs(_OutPath)
                    End If
                    ret = enmOutputReturn.OK
                Catch ex As Exception
                    ret = enmOutputReturn.ERR_SAVEAS
                End Try
            Else
                If _RemovalMacro Then
                    xlBook.SaveAs(_OutPath, Excel.XlFileFormat.xlOpenXMLWorkbook)
                Else
                    xlBook.SaveAs(_OutPath)
                End If
                ret = enmOutputReturn.OK
            End If
        End If

        If _ReportCheck Then
            Try
                xlBook.PrintOut()
                ret = enmOutputReturn.OK
            Catch ex As Runtime.InteropServices.COMException
                'PDF出力時　キャンセル対応
            Catch ex As Exception
                Throw ex
            End Try
        End If
        Return ret
    End Function

    ''' <summary>
    ''' 左ヘッダーに文字列を設定する
    ''' </summary>
    ''' <param name="sheet"></param>
    ''' <param name="font"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Protected Sub SetPageSetupLeftHeader(ByRef sheet As Excel.Worksheet, ByVal font As String, ByVal value As String)
        sheet.PageSetup.LeftHeader = "&""" & font & """" & value
    End Sub

    ''' <summary>
    ''' 右ヘッダーに文字列を設定する
    ''' </summary>
    ''' <param name="sheet"></param>
    ''' <param name="font"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Protected Sub SetPageSetupRightHeader(ByRef sheet As Excel.Worksheet, ByVal font As String, ByVal value As String)
        sheet.PageSetup.RightHeader = "&""" & font & """" & value
    End Sub

    ''' <summary>
    ''' 中央ヘッダーに文字列を設定する
    ''' </summary>
    ''' <param name="sheet"></param>
    ''' <param name="font"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Protected Sub SetPageSetupCenterHeader(ByRef sheet As Excel.Worksheet, ByVal font As String, ByVal value As String)
        sheet.PageSetup.CenterHeader = "&""" & font & """" & value
    End Sub

    ''' <summary>
    ''' 出力先ディレクトリ
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property OutDir() As String
        Get
            Return Me._OutDir
        End Get
    End Property

    ''' <summary>
    ''' 出力先パス
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Protected Property OutPath() As String
        Get
            Return Me._OutPath
        End Get
        Set(value As String)
            Me._OutPath = value
        End Set
    End Property

    ''' <summary>
    ''' Excel出力
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property ExcelCheck() As Boolean
        Get
            Return Me._ExcelCheck
        End Get
    End Property

    ''' <summary>
    ''' 用紙印刷
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property ReportCheck() As Boolean
        Get
            Return Me._ReportCheck
        End Get
    End Property

    ''' <summary>
    ''' 進捗加増
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Protected WriteOnly Property ProgressAddValue() As Integer
        Set(ByVal Value As Integer)
            If Not ProgressDialog Is Nothing Then
                ProgressDialog.AddValue = Value
            End If
        End Set
    End Property
End Class
