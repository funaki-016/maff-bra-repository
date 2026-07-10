''' <summary>
''' ComObjectプロセスクラス
''' </summary>
''' <remarks></remarks>
Public MustInherit Class ComObjectProcess

    ''' <summary>
    ''' ガベージコレクト強制
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub GCCollect()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    ''' <summary>
    ''' COMオブジェクト解放
    ''' </summary>
    ''' <param name="pObjCom"></param>
    ''' <remarks></remarks>
    Public Sub ReleaseComObject(Of T As Class)(ByRef pObjCom As T)
        Try
            If pObjCom Is Nothing Then
                Return
            End If
            If System.Runtime.InteropServices.Marshal.IsComObject(pObjCom) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pObjCom)
            End If
        Finally
            pObjCom = Nothing
        End Try
    End Sub
End Class
