Imports System.runtime.InteropServices

Public Class ClassRasOpt
    ''' <summary>
    ''' Changes the connection information for an entry in the phone book or creates a new phone-book entry.
    ''' </summary>
    ''' <param name="lpszPhonebook">Pointer to a null-terminated string that specifies the full path and file name of a phone-book (PBK) file. 
    ''' If this parameter is NULL, the function uses the current default phone-book file.</param>
    ''' <param name="lpszEntry">Pointer to a null-terminated string that specifies an entry name. 
    ''' If the entry name matches an existing entry, RasSetEntryProperties modifies the properties of that entry.
    ''' If the entry name does not match an existing entry, RasSetEntryProperties creates a new phone-book entry.</param>
    ''' <param name="lpRasEntry">Pointer to the RASENTRY structure that specifies the connection data to associate with the phone-book entry.</param>
    ''' <param name="dwEntryInfoSize">Specifies the size, in bytes, of the buffer identified by the lpRasEntry parameter.</param>
    ''' <param name="lpbDeviceInfo">This parameter is unused. The calling function should set this parameter to NULL.</param>
    ''' <param name="dwDeviceInfoSize">This parameter is unused. The calling function should set this parameter to zero.</param>
    ''' <returns></returns>
    <DllImport("rasapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Shared Function RasSetEntryProperties(ByVal lpszPhonebook As String, _
                                                 ByVal lpszEntry As String, _
                                                 ByRef lpRasEntry As RASENTRY, _
                                                 ByVal dwEntryInfoSize As Integer, _
                                                 ByVal lpbDeviceInfo As IntPtr, _
                                                 ByVal dwDeviceInfoSize As Integer) As UInt32
    End Function

    <DllImport("rasapi32.dll", CharSet:=CharSet.Auto)> _
   Public Shared Function RasGetEntryProperties(ByVal lpszPhoneBook As String, _
                         ByVal szEntry As String, _
                         ByVal lpbEntry As IntPtr, _
                         ByRef lpdwEntrySize As UInt32, _
                         ByVal lpb As IntPtr, _
                         ByVal lpdwSize As IntPtr) As UInt32
    End Function

    Public Shared Function GetPropValue(ByVal pro As RASEO, ByVal val As UInt32) As Boolean
        Try
            Dim x As UInt32 = val And pro
            If x <> 0 Then
                Return True
            Else
                Return False
            End If


        Catch ex As Exception

        End Try
    End Function
    Public Shared Function GetProp2Value(ByVal pro As RASEO2, ByVal val As UInt32) As Boolean
        Try
            Dim x As UInt32 = val And pro
            If x <> 0 Then
                Return True
            Else
                Return False
            End If


        Catch ex As Exception

        End Try
    End Function


    ''' <summary>
    ''' Defines the connection options for entries.
    ''' </summary>
    <Flags()> _
    Public Enum RASEO As UInteger
        ''' <summary>
        ''' No entry options specified.
        ''' </summary>
        None = &H0

        ''' <summary>
        ''' The country id, country code, and area code members are used to construct the phone number.
        ''' </summary>
        UseCountryAndAreaCodes = &H1

        ''' <summary>
        ''' The IP address specified by the entry will be used for the connection.
        ''' </summary>
        SpecificIPAddress = &H2

        ''' <summary>
        ''' The DNS addresses and WINS addresses specified by the entry will be used for the connection.
        ''' </summary>
        SpecificNameServers = &H4

        ''' <summary>
        ''' IP header compression will be used on PPP (Point-to-point) connections.
        ''' </summary>
        IPHeaderCompression = &H8

        ''' <summary>
        ''' The default route for IP packets is through the dial-up adapter while the connection is active.
        ''' </summary>
        RemoteDefaultGateway = &H10

        ''' <summary>
        ''' The remote access service (RAS) will disable the PPP LCP extensions.
        ''' </summary>
        DisableLcpExtensions = &H20

        ''' <summary>
        ''' The remote access service (RAS) displays a terminal window for user input before dialing the connection.
        ''' </summary>
        ''' <remarks>This member is only used when the entry is dialed by the component.</remarks>
        TerminalBeforeDial = &H40

        ''' <summary>
        ''' The remote access service displays a terminal window for user input after dialing the connection.
        ''' </summary>
        ''' <remarks>This member is only used when the entry is dialed by the component.</remarks>
        TerminalAfterDial = &H80

        ''' <summary>
        ''' The remote access service (RAS) will display a status monitor in the taskbar.
        ''' </summary>
        ModemLights = &H100

        ''' <summary>
        ''' The software compression will be negotiated by the link.
        ''' </summary>
        SoftwareCompression = &H200

        ''' <summary>
        ''' Only secure password schemes can be used to authenticate the client with the server.
        ''' </summary>
        RequireEncryptedPassword = &H400

        ''' <summary>
        ''' Only the Microsoft secure password scheme (MSCHAP) can be used to authenticate the client with the server.
        ''' </summary>
        RequireMSEncryptedPassword = &H800

        ''' <summary>
        ''' Data encryption must be negotiated successfully or the connection should be dropped.
        ''' </summary>
        ''' <remarks>This flag is ignored unless <see cref="RASEO.RequireMSEncryptedPassword"/> is also set.</remarks>
        RequireDataEncryption = &H1000

        ''' <summary>
        ''' The remote access service (RAS) logs on to the network after the point-to-point connection is established.
        ''' </summary>
        NetworkLogOn = &H2000

        ''' <summary>
        ''' The remote access service (RAS) uses the username, password, and domain of the currently logged on user when dialing this entry.
        ''' </summary>
        ''' <remarks>This flag is ignored unless the <see cref="RASEO.RequireMSEncryptedPassword"/> is also set.</remarks>
        UseLogOnCredentials = &H4000

        ''' <summary>
        ''' Indicates when an alternate phone number connects successfully, that number will become the primary phone number. 
        ''' </summary>
        PromoteAlternates = &H8000

        ''' <summary>
        ''' Check for an existing remote file system and remote printer bindings before making a connection to this phone book entry.
        ''' </summary>
        SecureLocalFiles = &H10000

        ''' <summary>
        ''' Indicates the Extensible Authentication Protocol (EAP) must be supported for authentication.
        ''' </summary>
        RequireEap = &H20000

        ''' <summary>
        ''' Indicates the Password Authentication Protocol (PAP) must be supported for authentication.
        ''' </summary>
        RequirePap = &H40000

        ''' <summary>
        ''' Indicates Shiva's Password Authentication Protocol (SPAP) must be supported for authentication.
        ''' </summary>
        RequireSpap = &H80000

        ''' <summary>
        ''' The connection will use custom encryption.
        ''' </summary>
        [Custom] = &H100000

        ''' <summary>
        ''' The remote access dialer should display the phone number being dialed.
        ''' </summary>
        PreviewPhoneNumber = &H200000

        ''' <summary>
        ''' Indicates all modems on the computer will share the same phone number.
        ''' </summary>
        SharedPhoneNumbers = &H800000

        ''' <summary>
        ''' The remote access dialer should display the username and password prior to dialing.
        ''' </summary>
        PreviewUserPassword = &H1000000

        ''' <summary>
        ''' The remote access dialer should display the domain name prior to dialing.
        ''' </summary>
        PreviewDomain = &H2000000

        ''' <summary>
        ''' The remote access dialer will display its progress while establishing the connection.
        ''' </summary>
        ShowDialingProgress = &H4000000

        ''' <summary>
        ''' Indicates the Challenge Handshake Authentication Protocol (CHAP) must be supported for authentication.
        ''' </summary>
        RequireChap = &H8000000

        ''' <summary>
        ''' Indicates the Challenge Handshake Authentication Protocol (CHAP) must be supported for authentication.
        ''' </summary>
        RequireMSChap = &H10000000

        ''' <summary>
        ''' Indicates the Challenge Handshake Authentication Protocol (CHAP) version 2 must be supported for authentication.
        ''' </summary>
        RequireMSChap2 = &H20000000

        ''' <summary>
        ''' Indicates MSCHAP must also send the LanManager hashed password.
        ''' </summary>
        ''' <remarks>This flag requires that <see cref="RASEO.RequireMSChap"/> must also be set.</remarks>
        RequireWin95MSChap = &H40000000

        ''' <summary>
        ''' The remote access service (RAS) must invoke a custom scripting assembly after establishing a connection to the server.
        ''' </summary>
        CustomScript = &H80000000UI
    End Enum


    ''' <summary>
    ''' Defines the additional connection options for entries.
    ''' </summary>
    <Flags()> _
    Public Enum RASEO2 As UInteger
        ''' <summary>
        ''' No additional entry options specified.
        ''' </summary>
        None = &H0

        ''' <summary>
        ''' Prevents remote users from using file and print servers over the connection.
        ''' </summary>
        SecureFileAndPrint = &H1

        ''' <summary>
        ''' Equivalent of clearing the Client for Microsoft Networks checkbox in the connection properties 
        ''' dialog box on the networking tab.
        ''' </summary>
        SecureClientForMSNet = &H2

        ''' <summary>
        ''' Changes the default behavior to not negotiate multilink.
        ''' </summary>
        DoNotNegotiateMultilink = &H4

        ''' <summary>
        ''' Use the default credentials to access network resources.
        ''' </summary>
        DoNotUseRasCredentials = &H8

        ''' <summary>
        ''' Use a pre-shared key for IPSec authentication.
        ''' </summary>
        ''' <remarks>This member is only used by L2TP/IPSec VPN connections.</remarks>
        UsePreSharedKey = &H10

        ''' <summary>
        ''' Indicates the connection is to the Internet.
        ''' </summary>
        Internet = &H20

        ''' <summary>
        ''' Disables NBT probing for this connection.
        ''' </summary>
        DisableNbtOverIP = &H40

        ''' <summary>
        ''' Ignore the device settings specified by the phone book entry.
        ''' </summary>
        UseGlobalDeviceSettings = &H80

        ''' <summary>
        ''' Automatically attempts to re-establish the connection if the connection is lost.
        ''' </summary>
        ReconnectIfDropped = &H100

        ''' <summary>
        ''' Use the same set of phone numbers for all subentries in a multilink connection.
        ''' </summary>
        SharePhoneNumbers = &H200


    End Enum


    Const MAX_PATH As Integer = 260

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

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
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
