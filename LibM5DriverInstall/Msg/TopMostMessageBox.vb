Public NotInheritable Class TopMostMessageBox
    Private Sub New()
    End Sub



    Public Shared Function Show(ByVal message As String) As DialogResult
        Return _Show(message, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.None, True)
    End Function

    Public Shared Function Show(ByVal message As String, ByVal title As String) As DialogResult
        Return _Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.None, True)
    End Function

    Public Shared Function Show(ByVal message As String, ByVal title As String, ByVal MessageBoxButtons As System.Windows.Forms.MessageBoxButtons) As DialogResult
        Return _Show(message, title, MessageBoxButtons, MessageBoxIcon.None, True)
    End Function


    Public Shared Function ShowNoBlock(ByVal message As String) As DialogResult
        Return ShowNoBlock(message, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.None)
    End Function

    Public Shared Function ShowNoBlock(ByVal message As String, ByVal title As String) As DialogResult
        Return ShowNoBlock(message, title, MessageBoxButtons.OK, MessageBoxIcon.None)
    End Function

    Public Shared Function ShowNoBlock(ByVal message As String, ByVal title As String, ByVal MessageBoxButtons As System.Windows.Forms.MessageBoxButtons) As DialogResult
        Return ShowNoBlock(message, title, MessageBoxButtons, MessageBoxIcon.None)
    End Function

    Public Shared Function ShowNologEmail(ByVal message As String, ByVal title As String, ByVal buttons As MessageBoxButtons, ByVal icon As System.Windows.Forms.MessageBoxIcon) As DialogResult
        Return _Show(message, title, buttons, icon, False)
    End Function

    Public Shared Function Show(ByVal message As String, ByVal title As String, ByVal buttons As MessageBoxButtons, ByVal icon As System.Windows.Forms.MessageBoxIcon) As DialogResult
        Return _Show(message, title, buttons, icon, True)
    End Function


    Class data
        Public message As String = ""
        Public title As String = ""
        Public buttons As MessageBoxButtons = MessageBoxButtons.OK
        Public icon As System.Windows.Forms.MessageBoxIcon = MessageBoxIcon.None
        Public LogEmail As Boolean = True
    End Class

    Public Shared Function ShowNoBlock(ByVal message As String, ByVal title As String, ByVal buttons As MessageBoxButtons, ByVal icon As System.Windows.Forms.MessageBoxIcon) As DialogResult

        Try
            Dim d As New data
            d.message = message
            d.title = title
            d.buttons = buttons
            d.icon = icon
            d.LogEmail = True

            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _showTh, d)


        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "TopMostMessageBox.Show", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Private Shared Sub _showTh(ByVal o As Object)
        Try
            Dim d As data = CType(o, data)

            _Show(d.message, d.title, d.buttons, d.icon, d.LogEmail)

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "TopMostMessageBox._showTh", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Shared Function _Show(ByVal message As String, ByVal title As String, ByVal buttons As MessageBoxButtons, _
                                  ByVal icon As System.Windows.Forms.MessageBoxIcon, ByVal logEmail As Boolean) As DialogResult
        ' Create a host form that is a TopMost window which will be the 
        ' parent of the MessageBox.

        Try

            Dim topmostForm As New Form()
            ' We do not want anyone to see this window so position it off the 
            ' visible screen and make it as small as possible
            topmostForm.Size = New System.Drawing.Size(1, 1)
            topmostForm.StartPosition = FormStartPosition.Manual
            Dim rect As System.Drawing.Rectangle = SystemInformation.VirtualScreen
            topmostForm.Location = New System.Drawing.Point(rect.Bottom + 10, rect.Right + 10)
            topmostForm.Show()
            ' Make this form the active form and make it TopMost
            topmostForm.Focus()
            topmostForm.BringToFront()
            topmostForm.TopMost = True
            ' If StatusLog IsNot Nothing Then StatusLog.print("MessageBox.Show: titel=""" + title + """, msg=""" + message.Replace(vbCrLf, "[CR] ") + """")

            '  If logEmail Then LibEmailServer.ClassLibraEmail.ClinetLibraSend("m.simonato@hemina.net", nLinea.ToString, "TopMessageBox: " + title, message, Now.ToString)

            ' Finally show the MessageBox with the form just created as its owner
            Dim result As DialogResult = MessageBox.Show(topmostForm, message, title, buttons, icon)
            topmostForm.Dispose()
            ' clean it up all the way
            Return result

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "TopMostMessageBox.Show", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    
End Class
