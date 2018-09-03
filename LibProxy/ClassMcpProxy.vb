
Imports dotras

Public Class ClassMcpProxy
    Public Const RegKeyConnections As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"

    Public Shared Function GetPhoneBookPathAllUser() As String
        Try
            Return RasPhoneBook.GetPhoneBookPath(RasPhoneBookType.AllUsers)
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetPhoneBookPathCurrentUser() As String
        Try
            Return RasPhoneBook.GetPhoneBookPath(RasPhoneBookType.User)
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function GetM5Connections() As List(Of String)

        Try
            Dim l As New List(Of String)


            Dim k As New LibProxy.RegUtils

            Dim conns() As String = My.Computer.Registry.CurrentUser.OpenSubKey(RegKeyConnections).GetValueNames()

            If conns IsNot Nothing AndAlso conns.Length > 0 Then

                For i As Integer = 0 To conns.Length - 1
                    If conns(i).ToUpper.StartsWith("M5_") Then

                        l.Add(conns(i))
                    End If
                Next

            End If

            Return l

        Catch ex As Exception

        End Try
    End Function

    Public Shared Function CreateOrCopyProxySettings(ByVal listConn As List(Of String)) As Boolean
        Try
            Const RegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
            Dim b() As Byte = DirectCast(My.Computer.Registry.CurrentUser.OpenSubKey(RegKey).GetValue("DefaultConnectionSettings"), Byte())

            If b IsNot Nothing Then

                For i As Integer = 0 To listConn.Count - 1
                    My.Computer.Registry.CurrentUser.OpenSubKey(RegKey, True).SetValue(listConn(i), b, Microsoft.Win32.RegistryValueKind.Binary)
                Next

            End If
            Return True
        Catch ex As Exception
            '  MessageBox.Show(ex.Message, "CopyProxySettings", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Shared Function GetRasEntries(ByVal RasPhoneBookPath As String) As List(Of String)

        Dim listRasEntry As New List(Of String)

        Try

            Dim pbk As RasPhoneBook = New RasPhoneBook()

            pbk.Open(RasPhoneBookPath)

            For Each entry As DotRas.RasEntry In pbk.Entries
                listRasEntry.Add(entry.Name)
            Next

        Catch ex As Exception

        End Try

        Return listRasEntry

    End Function



    Public Shared Sub ProxyCreate(ByVal rasPhoneBookPath As String)
        Try

            Dim listRasEntries As New List(Of String)

            listRasEntries = ClassMcpProxy.GetRasEntries(rasPhoneBookPath)

            ClassMcpProxy.CreateOrCopyProxySettings(listRasEntries)

        Catch ex As Exception

        End Try
    End Sub


End Class
