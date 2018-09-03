Imports System.Runtime.InteropServices
Imports DotRas.RasEntryOptions

Public Class ClassRasDll

    <DllImport("rasapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
Public Shared Function RasSetEntryProperties(ByVal lpszPhonebook As String, ByVal lpszEntry As String, ByRef lpRasEntry As RASENTRY, ByVal dwEntryInfoSize As Integer, _
                                             ByVal lpbDeviceInfo As IntPtr, ByVal dwDeviceInfoSize As Integer) As UInt32
    End Function

    <DllImport("rasapi32.dll", CharSet:=CharSet.Auto)> _
    Public Shared Function RasGetEntryProperties(ByVal lpszPhoneBook As String, ByVal szEntry As String, ByRef lpRasEntry As RASENTRY, ByRef lpdwEntrySize As UInt32, _
                          ByVal lpb As IntPtr, ByVal lpdwSize As IntPtr) As UInt32
    End Function

    '<DllImport("rasapi32.dll", CharSet:=CharSet.Auto)> _
    'Public Shared Function RasGetEntryProperties(ByVal lpszPhoneBook As String, _
    '                      ByVal szEntry As String, _
    '                      ByVal lpbEntry As IntPtr, _
    '                      ByRef lpdwEntrySize As UInt32, _
    '                      ByVal lpb As IntPtr, _
    '                      ByVal lpdwSize As IntPtr) As UInt32
    'End Function


    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
Public Structure RASENTRY
        Public dwSize As Integer
        Public dwfOptions As Integer
        Public dwCountryID As Integer
        Public dwCountryCode As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxAreaCode) + 1)> _
        Public szAreaCode As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxPhoneNumber) + 1)> _
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
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.MAX_PATH))> _
        Public szScript As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.MAX_PATH))> _
        Public szAutodialDll As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.MAX_PATH))> _
        Public szAutodialFunc As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxDeviceType) + 1)> _
        Public szDeviceType As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxDeviceName) + 1)> _
        Public szDeviceName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxPadType) + 1)> _
        Public szX25PadType As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxX25Address) + 1)> _
        Public szX25Address As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxFacilities) + 1)> _
        Public szX25Facilities As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxUserData) + 1)> _
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
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.MAX_PATH))> _
        Public szCustomDialDll As String
        Public dwVpnStrategy As Integer
        Public dwfOptions2 As Integer
        Public dwfOptions3 As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxDnsSuffix))> _
        Public szDnsSuffix As String
        Public dwTcpWindowSize As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.MAX_PATH))> _
        Public szPrerequisitePbk As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CInt(Ras.Internal.NativeMethods.RAS_MaxEntryName))> _
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

    Public Enum RasEntryOptions
        RASEO_UseCountrAndAreaCodes = &H1
        RASEO_SpecificIpAddr = &H2
        RASEO_SpecificNameServers = &H4
        RASEO_IpHeaderCompression = &H8
        RASEO_RemoteDefaultGateway = &H10
        RASEO_DisableLcpExtensions = &H20
        RASEO_TerminalBeforeDial = &H40
        RASEO_TerminalAfterDial = &H80
        RASEO_ModemLights = &H100
        RASEO_SwCompression = &H200
        RASEO_RequireEncrptedPw = &H400
        RASEO_RequireMsEncrptedPw = &H800
        RASEO_RequireDataEncrption = &H1000
        RASEO_NetworkLogon = &H2000
        RASEO_UseLogonCredentials = &H4000
        RASEO_PromoteAlternates = &H8000
        RASEO_SecureLocalFiles = &H10000
        RASEO_RequireEAP = &H20000
        RASEO_RequirePAP = &H40000
        RASEO_RequireSPAP = &H80000
        RASEO_Custom = &H100000
        RASEO_PreviewPhoneNumber = &H200000
        RASEO_SharedPhoneNumbers = &H800000
        RASEO_PreviewUserPw = &H1000000
        RASEO_PreviewDomain = &H2000000
        RASEO_ShowDialingProgress = &H4000000
        RASEO_RequireCHAP = &H8000000
        RASEO_RequireMsCHAP = &H10000000
        RASEO_RequireMsCHAP2 = &H20000000
        RASEO_RequireW95MSCHAP = &H40000000

    End Enum

    Public Enum RasEntryOptions2
        None = &H0
        SecureFileAndPrint = &H1
        SecureClientForMSNet = &H2
        DoNotNegotiateMultilink = &H4
        DoNotUseRasCredentials = &H8
        UsePreSharedKey = &H10
        Internet = &H20
        DisableNbtOverIP = &H40
        UseGlobalDeviceSettings = &H80
        ReconnectIfDropped = &H100
        SharePhoneNumbers = &H200
        SecureRoutingCompartment = &H400
        UseTypicalSettings = &H800
        IPv6SpecificNameServer = &H1000
        IPv6RemoteDefaultGateway = &H2000
        RegisterIPWithDns = &H4000
        UseDnsSuffixForRegistration = &H8000
        IPv4ExplicitMetric = &H10000
        IPv6ExplicitMetric = &H20000
        DisableIkeNameEkuCheck = &H40000
        DisableClassBasedStaticRoute = &H80000
        IPv6SpecificAddress = &H100000
        DisableMobility = &H200000
        RequireMachineCertificates = &H400000

    End Enum

    Public Enum RasEntryTypes
        RASET_Phone = 1
        RASET_Vpn = 2
        RASET_Direct = 3
        RASET_Internet = 4
    End Enum

    Public Enum RasEntryEncryption
        ET_None = 0
        ET_Require = 1
        ET_RequireMax = 2
        ET_Optional = 3
    End Enum

    Public Enum RasNetProtocols
        RASNP_NetBEUI = &H1
        RASNP_Ipx = &H2
        RASNP_Ip = &H4
        RASNP_Ipv6 = &H8
    End Enum

    Public Enum RasFramingProtocol
        RASFP_Ppp = &H1
        RASFP_Slip = &H2
        RASFP_Ras = &H4
    End Enum


End Class
