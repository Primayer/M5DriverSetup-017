Imports System.Threading

Public Class ClassMessageBoxOkNonBloccante

    Private Class classData
        Public message As String = ""
        Public Title As String = ""
        Public icon As System.Windows.Forms.MessageBoxIcon = MessageBoxIcon.None
    End Class

    Public Shared Sub Show(ByVal message As String, ByVal Title As String, ByVal icon As System.Windows.Forms.MessageBoxIcon)
        Try
            Dim p As New classData
            p.message = message
            p.Title = Title
            p.icon = icon
            ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf _msgboxErr), p)
        Catch ex As Exception

        End Try
    End Sub

    Private Shared Sub _msgboxErr(ByVal p As Object)
        Try
            Dim ms As classData = DirectCast(p, classData)
            'MessageBox.Show(ms.message, ms.Title, MessageBoxButtons.OK, ms.icon)

            TopMostMessageBox.Show(ms.message, ms.Title, MessageBoxButtons.OK, ms.icon)
           
        Catch ex As Exception
        End Try

    End Sub

   
End Class

