Option Explicit On
Option Strict On

Public Class CatiaDataCOMCleaner

    Public Sub Release(ParamArray objects As Object())
        If objects Is Nothing Then Exit Sub

        For Each obj In objects
            If obj IsNot Nothing AndAlso Runtime.InteropServices.Marshal.IsComObject(obj) Then
                Try
                    Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
                Catch
                    ' Objeto ya liberado o inválido
                End Try
            End If
        Next

        ' Indispensable para que los procesos externos (CNEXT/EXCEL) se cierren
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub
End Class