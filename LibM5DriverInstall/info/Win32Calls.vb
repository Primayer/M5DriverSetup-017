Imports System.Runtime.InteropServices

Class Win32Calls
    Public Shared ANYSIZE_ARRAY As UInteger = 1000

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SP_DEVINFO_DATA
        Public cbSize As Integer
        Public ClassGuid As Guid
        Public DevInst As UInteger
        Public Reserved As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SP_DEVICE_INTERFACE_DATA
        Public cbSize As Integer
        Public InterfaceClassGuid As Guid
        Public Flags As UInteger
        Public Reserved As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SP_DEVICE_INTERFACE_DETAIL_DATA
        Public cbSize As Integer
        Public DevicePath As Byte()
    End Structure

    <DllImport("Kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetLastError() As UInteger
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiGetClassDevs(ByRef ClassGuid As Guid, ByVal Enumerator As Integer, ByVal hwndParent As IntPtr, ByVal Flags As UInt32) As IntPtr
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiEnumDeviceInterfaces(ByVal hDevInfo As IntPtr, ByVal devInfo As IntPtr, ByRef interfaceClassGuid As Guid, ByVal memberIndex As UInt32, ByRef deviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As [Boolean]
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiGetDeviceInterfaceDetail(ByVal nSetupDiGetClassDevs As IntPtr, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByRef DeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA, ByVal DeviceInterfaceDetailDataSize As UInteger, ByRef RequiredSize As Integer, ByRef DeviceInfoData As SP_DEVINFO_DATA) As [Boolean]
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiGetDeviceInterfaceDetail(ByVal nSetupDiGetClassDevs As IntPtr, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal Ptr As IntPtr, ByVal DeviceInterfaceDetailDataSize As UInteger, ByRef RequiredSize As Integer, ByVal PtrInfo As IntPtr) As [Boolean]
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CM_Get_Parent(ByRef pdnDevInst As UInt32, ByVal dnDevInst As UInteger, ByVal ulFlags As ULong) As UInteger
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto)> _
    Public Shared Function CM_Get_Device_ID(ByVal dnDevInst As UInt32, ByVal Buffer As IntPtr, ByVal BufferLen As Integer, ByVal ulFlags As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiDestroyDeviceInfoList(ByVal hDevInfo As IntPtr) As Boolean
    End Function
    <DllImport("hid.dll", SetLastError:=True)> _
    Public Shared Sub HidD_GetHidGuid(ByRef gHid As Guid)
    End Sub
End Class
