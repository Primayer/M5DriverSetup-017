'Imports System.ComponentModel
'Imports System.Runtime.InteropServices
'Imports System.Management
'' This code demonstrates how to monitor the UsbControllerDevice for the arrival of pnp device creation/operation events
'Class WMIEvent
'    Public Shared Sub Main()
'        Dim usbEvnt As New USBEvent()
'        usbEvnt.Start()
'        Console.ReadLine()
'        ' block main thread for test purposes
'        usbEvnt.[Stop]()
'    End Sub

'End Class

'Friend Class USBEvent

'    Dim w As ManagementEventWatcher = Nothing

'    Friend Sub Start()
'        ' Bind to local machine

'        Try
'            Dim q As New WqlEventQuery()
'            q.EventClassName = "__InstanceOperationEvent"
'            q.WithinInterval = New TimeSpan(0, 0, 3)
'            q.Condition = "TargetInstance ISA 'Win32_USBControllerDevice' "
'            w = New ManagementEventWatcher(q)
'            ' w.EventArrived += New EventArrivedEventHandler(Me.UsbEventArrived)
'            AddHandler w.EventArrived, New EventArrivedEventHandler(AddressOf Me.UsbEventArrived)
'            w.Start()
'            ' Start listen for events
'        Catch ex As Exception

'        End Try

'    End Sub

'    Friend Sub [Stop]()
'        '  w.EventArrived -= New EventArrivedEventHandler(Me.UsbEventArrived)
'        RemoveHandler w.EventArrived, New EventArrivedEventHandler(AddressOf Me.UsbEventArrived)
'        w.[Stop]()
'        w.Dispose()
'    End Sub

'    Public Sub UsbEventArrived(ByVal sender As Object, ByVal e As EventArrivedEventArgs)
'        'Get the Event object and display all it properties
'        Dim mbo As ManagementBaseObject = DirectCast(e.NewEvent("TargetInstance"), ManagementBaseObject)
'        Using o As New ManagementObject(mbo("Dependent").ToString())
'            Dim s As String = ""
'            o.[Get]()
'            If o.Properties.Count = 0 Then
'                s &= "No further device properties, device removed?"
'            End If
'            For Each prop As PropertyData In o.Properties
'                s &= String.Format("{0} - {1}", prop.Name, prop.Value) + vbCrLf
'            Next
'            s &= "-------------------------------"
'            MessageBox.Show(s)
'        End Using

'    End Sub

'End Class

