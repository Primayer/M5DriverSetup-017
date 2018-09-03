Public Class ClassShowHideDevice
    Public Shared Sub DevMgr_Show_NonPresent_Devices(ByVal ShowHide As Boolean)
        Try
            Dim variable As String = "DevMgr_Show_NonPresent_Devices"
            Dim value As String = "0"

            If ShowHide Then
                value = "1"
            End If

            Environment.SetEnvironmentVariable(variable, value)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
