Option Strict Off

Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop.Forms

''' <summary>
''' 共通処理(遅延バインディング専用)
''' </summary>
''' <remarks></remarks>
Public Class ComUtilStrictOff

    ''' <summary>
    ''' Excelのユーザーフォームを取得する
    ''' </summary>
    ''' <param name="pXlBook"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelForm(ByVal pXlBook As Excel.Workbook) As UserForm
        Return DirectCast(pXlBook.Form, UserForm)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「保存して前画面」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnSaveClose(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnSaveClose, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「保存せず前画面」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnNoSaveClose(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnNoSaveClose, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「再計算」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnCalculate(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnCalculate, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「基本エラーチェック」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnBaseErrorCheck(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnBaseErrorCheck, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「範囲エラーチェック」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnRangeErrorCheck(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnRangeErrorCheck, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「追加エラーチェック」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnAddErrorCheck(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnAddErrorCheck, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「エラーチェック一覧」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnErrorCheckList(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnErrorCheckList, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「保存（再計算あり）」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnSaveCalculate(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnSaveCalculate, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「保存（再計算なし）」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnSaveNoCalculate(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnSaveNoCalculate, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「修正前データ表示」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnBeforeDataShow(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnBeforeDataShow, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「ファイル入力」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnFileInput(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnFileInput, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームから「ファイル出力」ボタンを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExcelBtnFileOutput(ByVal pXlForm As UserForm) As CommandButton
        Return DirectCast(pXlForm.btnFileOutput, CommandButton)
    End Function

    ''' <summary>
    ''' Excelのユーザーフォームを表示する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowExcelForm(ByVal pXlForm As userform)
        pXlForm.Show()
    End Sub

    ''' <summary>
    ''' Excelのユーザーフォームのハンドルを取得する
    ''' </summary>
    ''' <param name="pXlForm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFormHwnd(ByVal pXlForm As UserForm) As Integer
        Return pXlForm.hWnd
    End Function
End Class