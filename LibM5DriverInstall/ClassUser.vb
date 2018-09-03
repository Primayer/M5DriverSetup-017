

Public Class ClassUser

    Public Shared Function GetName() As String
        Try
            Return My.User.CurrentPrincipal.Identity.Name

        Catch ex As Exception

        End Try
    End Function

    Public Shared Function IsAdministrator() As Boolean
        Try
            If My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Shared Function msgControlPrivileges() As Boolean
        Try
            If Not ClassUser.IsAdministrator Then
                MessageBox.Show("Only with administrator privileges you can update power tool!", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            Else
                Return True
            End If

        Catch ex As Exception

        End Try
    End Function

End Class
