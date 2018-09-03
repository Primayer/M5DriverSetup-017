Imports System.Runtime.InteropServices
Public Class ClassRasEntryProperties


    <DllImport("rasapi32.dll", CharSet:=CharSet.Auto)> _
    Public Shared Function RasGetEntryProperties(ByVal lpszPhoneBook As String, _
                          ByVal szEntry As String, _
                          ByVal lpbEntry As IntPtr, _
                          ByRef lpdwEntrySize As UInt32, _
                          ByVal lpb As IntPtr, _
                          ByVal lpdwSize As IntPtr) As UInt32
    End Function

    Const MAX_PATH As Integer = 260

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure RASENTRY
        Public dwSize As Integer
        Public dwfOptions As Integer
        Public dwCountryID As Integer
        Public dwCountryCode As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxAreaCode) + 1)> _
        Public szAreaCode As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxPhoneNumber) + 1)> _
        Public szLocalPhoneNumber As String
        Public dwAlternateOffset As Integer
        Public ipaddr As RASIPADDR
        Public ipaddrDns As RASIPADDR
        Public ipaddrDnsAlt As RASIPADDR
        Public ipaddrWins As RASIPADDR
        Public ipaddrWinsAlt As RASIPADDR
        Public dwFrameSize As Integer
        Public dwfNetProtocols As Integer
        Public dwFramingProtocol As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(MAX_PATH))> _
        Public szScript As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(MAX_PATH))> _
        Public szAutodialDll As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(MAX_PATH))> _
        Public szAutodialFunc As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxDeviceType) + 1)> _
        Public szDeviceType As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxDeviceName) + 1)> _
        Public szDeviceName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxPadType) + 1)> _
        Public szX25PadType As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxX25Address) + 1)> _
        Public szX25Address As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxFacilities) + 1)> _
        Public szX25Facilities As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxUserData) + 1)> _
        Public szX25UserData As String
        Public dwChannels As Integer
        Public dwReserved1 As Integer
        Public dwReserved2 As Integer
        Public dwSubEntries As Integer
        Public dwDialMode As Integer
        Public dwDialExtraPercent As Integer
        Public dwDialExtraSampleSeconds As Integer
        Public dwHangUpExtraPercent As Integer
        Public dwHangUpExtraSampleSeconds As Integer
        Public dwIdleDisconnectSeconds As Integer
        Public dwType As Integer
        Public dwEncryptionType As Integer
        Public dwCustomAuthKey As Integer
        Public guidId As Guid
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(MAX_PATH))> _
        Public szCustomDialDll As String
        Public dwVpnStrategy As Integer
        Public dwfOptions2 As Integer
        Public dwfOptions3 As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxDnsSuffix))> _
        Public szDnsSuffix As String
        Public dwTcpWindowSize As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(MAX_PATH))> _
        Public szPrerequisitePbk As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(RasFieldSizeConstants.RAS_MaxEntryName))> _
        Public szPrerequisiteEntry As String
        Public dwRedialCount As Integer
        Public dwRedialPause As Integer
        Private ipv6addrDns As RASIPV6ADDR
        Private ipv6addrDnsAlt As RASIPV6ADDR
        Public dwIPv4InterfaceMetric As Integer
        Public dwIPv6InterfaceMetric As Integer
        Private ipv6addr As RASIPV6ADDR
        Public dwIPv6PrefixLength As Integer
        Public dwNetworkOutageTime As Integer
    End Structure

    Public Enum RasFieldSizeConstants
        RAS_MaxDeviceType = 16
        RAS_MaxPhoneNumber = 128
        RAS_MaxIpAddress = 15
        RAS_MaxIpxAddress = 21
        RAS_MaxEntryName = 256
        RAS_MaxDeviceName = 128
        RAS_MaxCallbackNumber = RAS_MaxPhoneNumber
        RAS_MaxAreaCode = 10
        RAS_MaxPadType = 32
        RAS_MaxX25Address = 200
        RAS_MaxFacilities = 200
        RAS_MaxUserData = 200
        RAS_MaxReplyMessage = 1024
        RAS_MaxDnsSuffix = 256
        UNLEN = 256
        PWLEN = 256
        DNLEN = 15
    End Enum

    Public Structure RASIPADDR
        Private a As Byte
        Private b As Byte
        Private c As Byte
        Private d As Byte
    End Structure

    Public Structure RASIPV6ADDR
        Private a As Byte
        Private b As Byte
        Private c As Byte
        Private d As Byte
        Private e As Byte
        Private f As Byte
    End Structure

End Class
