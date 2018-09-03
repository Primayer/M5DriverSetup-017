Public Class ClassMsgErr
    Public Shared Sub LogErr(ByVal [class] As String, ByVal [function] As String, ByVal ex As Exception, ByVal show As Boolean)
        Try
            If show Then
                MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace + vbCrLf + vbCrLf + "Class: " + [class], "Function: " + [function], MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex1 As Exception

        End Try
    End Sub


End Class
