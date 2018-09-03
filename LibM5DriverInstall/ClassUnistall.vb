Imports System.Runtime.InteropServices
Imports System.ComponentModel

Imports System.Text


Public Class ClassUnistall
    Public Const CR_SUCCESS As Integer = &H0
    Public Const CM_LOCATE_DEVNODE_NORMAL As Integer = &H0
    Public Const DIGCF_PRESENT As Integer = &H2
    Public Const DIF_REMOVE As Integer = &H5

    Const ERROR_SUCCESS As Integer = 0
    Const DICS_FLAG_GLOBAL As Int32 = &H1
    Const MAX__DESC As Integer = 256
    Const DIREG_DEV As Integer = &H1
    Const KEY_READ As Integer = &H20019

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
Public Structure SP_DEVINFO_DATA
        Public cbSize As Integer
        Public classGuid As Guid
        Public propertyId As Integer
        Public reserved As IntPtr
    End Structure

    ' importing external SetupAPI methods
    <DllImport("cfgmgr32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CM_Locate_DevNode(ByRef DevInst As UInt32, ByVal pDeviceID As String, ByVal Flags As UInt32) As UInt32
    End Function

    <DllImport("cfgmgr32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CM_Reenumerate_DevNode(ByVal DevInst As UInt32, ByVal Flags As UInt32) As UInt32
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiClassGuidsFromNameA(ByVal ClassName As String, ByRef Guids As Guid, ByVal ClassNameSize As UInt32, ByRef RequiredSize As UInt32) As [Boolean]
    End Function
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
        Private Shared Function SetupDiClassGuidsFromName( _
        ByVal ClassName As StringBuilder, _
        ByRef ClassGuids As Guid, _
        ByVal ClassGuidSize As Integer, _
        ByRef ClassGuidRequiredSize As Integer) As Boolean
    End Function

    <DllImport("setupapi.dll", EntryPoint:="SetupDiGetClassDevsW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=True, CallingConvention:=CallingConvention.Winapi)> _
  Private Shared Function SetupDiGetClassDevs( _
      ByRef ClassGuid As Guid, _
      ByVal Enumerator As Integer, _
      ByVal hwndParent As Integer, _
      ByVal Flags As Integer) As IntPtr
    End Function


    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiGetClassDevsA(ByRef ClassGuid As Guid, ByVal Enumerator As UInt32, ByVal hwndPtr As IntPtr, ByVal Flags As UInt32) As IntPtr
    End Function

    '<DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    'Public Shared Function SetupDiEnumDeviceInfo(ByVal DeviceInfoSet As IntPtr, ByVal DeviceIndex As UInt32, ByVal DeviceInfoData As SP_DEVINFO_DATA) As [Boolean]
    'End Function

    <DllImport("setupapi.dll", EntryPoint:="SetupDiEnumDeviceInfo", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=True, CallingConvention:=CallingConvention.Winapi)> _
      Private Shared Function SetupDiEnumDeviceInfo( _
          ByVal DeviceInfoSet As Integer, _
          ByVal MemberIndex As Integer, _
          ByRef DeviceInfoData As SP_DEVINFO_DATA) As Boolean
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiCallClassInstaller(ByVal InstallFunction As UInt32, ByVal DeviceInfoSet As IntPtr, ByRef DeviceInfoData As SP_DEVINFO_DATA) As [Boolean]
    End Function
     

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetupDiDestroyDeviceInfoList(ByVal DeviceInfoSet As IntPtr) As [Boolean]
    End Function

    <DllImport("Setupapi", CharSet:=CharSet.Auto, SetLastError:=True)> _
   Public Shared Function SetupDiOpenDevRegKey(ByVal hDeviceInfoSet As IntPtr, ByRef deviceInfoData As SP_DEVINFO_DATA, ByVal scope As Integer, ByVal hwProfile As Integer, _
                                               ByVal parameterRegistryValueKind As Integer, ByVal samDesired As Integer) As IntPtr
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
       Private Shared Function RegCloseKey(ByVal hKey As IntPtr) As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
   Friend Shared Function RegQueryValueEx(ByVal hKey As IntPtr, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As System.Text.StringBuilder, ByRef lpcbData As Integer) As Integer
    End Function

    Public Sub Example()
        Try
            Dim ClassNameB As New StringBuilder("Modem")
            'Dim ClassName As New StringBuilder("net")
            Dim ClassGuid As Guid
            Dim GuidSize As Integer = 0
            Dim GuidReqtSize As Integer
            Dim intRtrn As Boolean = False
            intRtrn = SetupDiClassGuidsFromName(ClassNameB, ClassGuid, GuidSize, GuidReqtSize)
            GuidSize = GuidReqtSize
            intRtrn = SetupDiClassGuidsFromName(ClassNameB, ClassGuid, GuidSize, GuidReqtSize)
            MsgBox(ClassGuid.ToString)

        Catch ex As Exception

        End Try
       
    End Sub

    Public Sub DisplayError(ByVal title As String, ByVal desc As String)
        Try
            Dim msg As String = ""
            Dim errore As Integer = Marshal.GetLastWin32Error
            msg = "Error: " + errore.ToString + " : " + New Win32Exception(errore).Message + vbCrLf + vbCrLf + desc
            MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "DisplayError", ex, True)
        End Try
    End Sub

    Public Function RemoveAllDevicesModem() As Boolean

        Try
            Dim devInfoElem As SP_DEVINFO_DATA
            Dim index As Integer = 0
            Dim esito As Boolean = True

            Dim CLASS_GUID_MODEM As Guid = New Guid("{4D36E96D-E325-11CE-BFC1-08002BE10318}") ' modem

            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_MODEM, 0, 0, DIGCF_PRESENT)

            devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

            While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True
                ' invoke class installer method with DIF_REMOVE flag
                If Not SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, devInfoElem) Then
                    DisplayError("SetupDiCallClassInstaller", "")
                    ' failed to uninstall device
                    SetupDiDestroyDeviceInfoList(hDeviceInfoSet)
                    esito = False
                    Exit While
                End If
                index = index + 1
            End While

            ' perform cleanup
            SetupDiDestroyDeviceInfoList(hDeviceInfoSet)

            Return esito

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassUnistall", "RemoveAllDevicesModem", ex, True)
            Return False
        End Try

    End Function

    Const ERROR_INSUFFICIENT_BUFFER As Integer = 122

    <DllImport("setupapi.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
Private Shared Function SetupDiGetDeviceInterfaceDetail(ByVal deviceInfoSet As IntPtr, ByRef deviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal deviceInterfaceDetailData As IntPtr, ByVal deviceInterfaceDetailDataSize As Integer, ByRef requiredSize As Integer, ByRef deviceInfoData As SP_DEVINFO_DATA) As Boolean
    End Function

    Public Declare Auto Function SetupDiGetDeviceInterfaceDetail2 Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailW" ( _
    ByVal hDevInfo As IntPtr, _
    ByRef deviceInterfaceData As SP_DEVICE_INTERFACE_DATA, _
    ByRef deviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA, _
    ByVal deviceInterfaceDetailDataSize As Int32, _
    ByRef requiredSize As Int32, _
    ByRef deviceInfoData As SP_DEVINFO_DATA) As Boolean

    <StructLayout(LayoutKind.Sequential)> _
Public Structure SP_DEVICE_INTERFACE_DETAIL_DATA
        Public cbSize As UInt32
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> _
        Public DevicePath As String
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
Class SP_DEVICE_INTERFACE_DATA
        Public cbSize As Integer = Marshal.SizeOf(GetType(SP_DEVICE_INTERFACE_DATA))
        Public interfaceClassGuid As Guid = Guid.Empty
        ' temp
        Public flags As Integer = 0
        Public reserved As Integer = 0
    End Class

    Public Function RemoveDevicesModem(ByVal PortCom As String) As Boolean

        Try
            Dim devInfoElem As SP_DEVINFO_DATA
            Dim index As Integer = 0
            Dim esito As Boolean = True
            Dim comPortx As String = ""

            Dim CLASS_GUID_MODEM As Guid = New Guid("{4D36E96D-E325-11CE-BFC1-08002BE10318}") ' modem
            '{4D36E96D-E325-11CE-BFC1-08002BE10318}

            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_MODEM, 0, 0, DIGCF_PRESENT)

            devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

            While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True

                Dim interfaceData As New SP_DEVICE_INTERFACE_DATA
                interfaceData.cbSize = Marshal.SizeOf(interfaceData)
               
                Dim size As Integer = 0

                Dim dd As New SP_DEVICE_INTERFACE_DETAIL_DATA
                dd.cbSize = CUInt(Marshal.SizeOf(GetType(SP_DEVICE_INTERFACE_DETAIL_DATA)))

                If Not SetupDiGetDeviceInterfaceDetail2(hDeviceInfoSet, interfaceData, dd, 0, size, devInfoElem) Then
                    Dim errore As Integer = Marshal.GetLastWin32Error
                    If Not errore = ERROR_INSUFFICIENT_BUFFER Then
                        Throw New Win32Exception(errore)
                    End If

                End If


                If 1 = 0 Then
                    ' invoke class installer method with DIF_REMOVE flag
                    If Not SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, devInfoElem) Then
                        DisplayError("SetupDiCallClassInstaller", "")
                        ' failed to uninstall device
                        SetupDiDestroyDeviceInfoList(hDeviceInfoSet)
                        esito = False
                        Exit While
                    End If
                End If
                index = index + 1
            End While

            ' perform cleanup
            SetupDiDestroyDeviceInfoList(hDeviceInfoSet)

            Return esito

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassUnistall", "RemoveAllDevicesModem", ex, True)
            Return False
        End Try

    End Function


    Public Function RemoveDevicesModem1(ByVal PortCom As String) As Boolean

        Try
            Dim devInfoElem As SP_DEVINFO_DATA
            Dim index As Integer = 0
            Dim esito As Boolean = True
            Dim comPortx As String = ""

            Dim CLASS_GUID_MODEM As Guid = New Guid("{4D36E96D-E325-11CE-BFC1-08002BE10318}") ' modem
            '{4D36E96D-E325-11CE-BFC1-08002BE10318}

            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_MODEM, 0, 0, DIGCF_PRESENT)

            devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

            While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True

                Dim hKeyDev As System.IntPtr

                ' Opens a registry key for device-specific configuration information.
                hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ)
                Dim szDevDesc As New System.Text.StringBuilder(20, MAX__DESC)

                If hKeyDev <> New IntPtr(-1) AndAlso hDeviceInfoSet <> New IntPtr(-1) AndAlso hDeviceInfoSet <> IntPtr.Zero Then
                    Dim esitor As Integer = RegQueryValueEx(hKeyDev, "HardwareID", Nothing, Nothing, szDevDesc, MAX__DESC)
                    If (ERROR_SUCCESS = esitor) Then
                        RegCloseKey(hKeyDev)
                        Dim DevDesc As String = szDevDesc.ToString
                    End If
                End If

                If PortCom = comPortx Then
                    ' invoke class installer method with DIF_REMOVE flag
                    If Not SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, devInfoElem) Then
                        DisplayError("SetupDiCallClassInstaller", "")
                        ' failed to uninstall device
                        SetupDiDestroyDeviceInfoList(hDeviceInfoSet)
                        esito = False
                        Exit While
                    End If
                End If
                index = index + 1
            End While

            ' perform cleanup
            SetupDiDestroyDeviceInfoList(hDeviceInfoSet)

            Return esito

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassUnistall", "RemoveAllDevicesModem", ex, True)
            Return False
        End Try

    End Function

    Private Function ScanForHardwareChanges() As Int32
        Dim DevInst As UInt32 = 0
        Dim result As UInt32 = CM_Locate_DevNode(DevInst, Nothing, CM_LOCATE_DEVNODE_NORMAL)
        If result <> CR_SUCCESS Then
            ' failed to get devnode
            Return -1
        End If
        result = CM_Reenumerate_DevNode(DevInst, 0)

        If result <> CR_SUCCESS Then
            ' reenumeration failed
            Return -2
        End If
        Return CR_SUCCESS
    End Function

End Class
