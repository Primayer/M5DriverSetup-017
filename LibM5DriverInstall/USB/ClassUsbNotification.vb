Imports System.Runtime.InteropServices

Public NotInheritable Class UsbNotification

    ' device change event      
    Public Const DBT_DEVTYP_DEVICEINTERFACE As Integer = &H5
    Public Const DBT_DEVTYP_PORT As Integer = &H3



    ' There has been a change in the devices
    Public Const WM_DEVICECHANGE As Integer = &H219


    ' System detects a new device: WM_DEVICECHANGE message
    Public Const DBT_DEVICEARRIVAL As Integer = &H8000
    ' Device removal request
    Public Const DBT_DEVICEQUERYREMOVE As Integer = &H8001
    ' Device removal failed
    Public Const DBT_DEVICEQUERYREMOVEFAILED As Integer = &H8002
    ' Device removal is pending
    Public Const DBT_DEVICEREMOVEPENDING As Integer = &H8003
    ' The device has been succesfully removed from the system
    Public Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    ' Logical Volume (A disk has been inserted, such a usb key or external HDD)
    Public Const DBT_DEVTYP_VOLUME As Integer = &H2

  
    Structure DEV_BROADCAST_HDR
        Public dbch_Size As UInteger
        Public dbch_DeviceType As UInteger
        Public dbch_Reserved As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Structure DEV_BROADCAST_PORT
        Public dbcp_size As Integer
        Public dbcp_devicetype As Integer
        Public dbcp_reserved As Integer
        ' MSDN say "do not use"
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=255)> _
        Public dbcp_name As [String]
    End Structure


    'Private Shared ReadOnly GuidDevinterfaceUSBDevice As New Guid("A5DCBF10-6530-11D2-901F-00C04FB951ED")

    ' Ports
    ' Private Shared ReadOnly GuidDevinterfaceUSBDevice As New Guid("{4d36e978-e325-11ce-bfc1-08002be10318}")


    ' USB Raw Device Interface Class GUID
    'Private Shared ReadOnly GuidDevinterfaceUSBDevice As New Guid(&HA5DCBF10, &H6530, &H11D2, &H90, &H1F, &H0, &HC0, &H4F, &HB9, &H51, &HED)

    ' Disk Device Interface Class GUID
    '{ 0x53f56307, 0xb6bf, 0x11d0, { 0x94, 0xf2, 0x00, 0xa0, 0xc9, 0x1e, 0xfb,0x8b } },

    'Human Interface Device Class GUID
    '{ 0x4d1e55b2, 0xf16f, 0x11Cf, { 0x88, 0xcb, 0x00, 0x11, 0x11, 0x00, 0x00,0x30 } },

    ' FTDI_D2XX_Device Class GUID
    '{ 0x219d0508, 0x57a8, 0x4ff5, {0x97, 0xa1, 0xbd, 0x86, 0x58, 0x7c, 0x6c,0x7e } },

    ' FTDI_VCP_Device Class GUID
    Private Shared ReadOnly GuidDevinterfaceUSBDevice As New Guid(&H86E0D1E0L, &H8089, &H11D0, &H9C, &HE4, &H8, &H0, &H3E, &H30, &H1F, &H73)


    ' USB devices
    Private Shared notificationHandle As IntPtr

    ''' <summary>
    ''' Registers a window to receive notifications when USB devices are plugged or unplugged.
    ''' </summary>
    ''' <param name="windowHandle">Handle to the window receiving notifications.</param>
    Public Shared Sub RegisterUsbDeviceNotification(ByVal windowHandle As IntPtr)
        Dim dbi As New DevBroadcastDeviceinterface
        dbi.DeviceType = DBT_DEVTYP_DEVICEINTERFACE
        dbi.Reserved = 0
        dbi.ClassGuid = GuidDevinterfaceUSBDevice
        dbi.Name = 0

        dbi.Size = Marshal.SizeOf(dbi)
        Dim buffer As IntPtr = Marshal.AllocHGlobal(dbi.Size)
        Marshal.StructureToPtr(dbi, buffer, True)

        notificationHandle = RegisterDeviceNotification(windowHandle, buffer, 0)
    End Sub

    ''' <summary>
    ''' Unregisters the window for USB device notifications
    ''' </summary>
    Public Shared Sub UnregisterUsbDeviceNotification()
        UnregisterDeviceNotification(notificationHandle)
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function RegisterDeviceNotification(ByVal recipient As IntPtr, ByVal notificationFilter As IntPtr, ByVal flags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function UnregisterDeviceNotification(ByVal handle As IntPtr) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure DevBroadcastDeviceinterface
        Friend Size As Integer
        Friend DeviceType As Integer
        Friend Reserved As Integer
        Friend ClassGuid As Guid
        Friend Name As Short
    End Structure


End Class