Public Class ClassDevice
    ' class guid for FTDI: Class Ports
    ' Device manager, right click a COM PORT (that can be disabled), click properties, 
    ' Details tab, select "device class guid" property:
    Private Shared PortsGuid As New Guid("{4d36e978-e325-11ce-bfc1-08002be10318}")

    ' instance id for some weird infrared mouse I have. "device instance path" in device manager:
    ' Private Shared instanceId As String = "FTDIBUS\VID_0403+PID_6015+6&2DB47F2C&0&3\0000"

    'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\FTDIBUS\VID_0403+PID_6015+6&2db47f2c&0&3\0000

    Public Shared Function Enable(ByVal ComPort As String) As Boolean
        Return EnableDisable(ComPort, True)
    End Function

    Public Shared Function Disable(ByVal ComPort As String) As Boolean
        Return EnableDisable(ComPort, False)
    End Function

    Private Shared Function EnableDisable(ByVal ComPort As String, ByVal OnOff As Boolean) As Boolean
        Try
            Dim instanceId As String = ""

            instanceId = GetIstanceID(ComPort)

            If instanceId = "" Then
                Return False
            End If

            Return DeviceHelper.SetDeviceEnabled(PortsGuid, instanceId, OnOff)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassDevice", "EnableDisable", ex, True)
        End Try
    End Function

    Public Shared Function GetIstanceID(ByVal ComPort As String) As String

        Try
            Dim listDeviceID As New List(Of ClassFtdiRegKey.ClassDeviceID)

            listDeviceID = ClassFtdiRegKey.SearchFtdiDeviceID

            For i As Integer = 0 To listDeviceID.Count - 1
                If ComPort.ToUpper = listDeviceID(i).PortName Then
                    Return listDeviceID(i).DeviceID
                End If
            Next
            Return ""
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassDevice", "GetIstanceID", ex, True)
            Return ""
        End Try
    End Function

    Public Shared Sub GetIstanceIDList(ByVal lstbox As ListBox)

        Try
            If lstbox IsNot Nothing Then lstbox.Items.Clear()

            Dim listDeviceID As New List(Of ClassFtdiRegKey.ClassDeviceID)

            listDeviceID = ClassFtdiRegKey.SearchFtdiDeviceID()

            If lstbox IsNot Nothing Then
                For i As Integer = 0 To listDeviceID.Count - 1
                    lstbox.Items.Add(listDeviceID(i).PortName + ": " + listDeviceID(i).DeviceID)
                Next
            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassDevice", "GetIstanceIDList", ex, True)
        End Try
    End Sub
End Class
