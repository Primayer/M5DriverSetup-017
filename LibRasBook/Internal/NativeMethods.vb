Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME

Imports DotRas

Namespace Ras.Internal

    ''' <summary>
    ''' Contains the remote access service (RAS) API entry points and structure definitions.
    ''' </summary>
    Friend NotInheritable Class NativeMethods
        Private Sub New()
        End Sub
#Region "Fields"

        ''' <summary>
        ''' Defines the default polling interval, in milliseconds, used when disconnecting an active connection.
        ''' </summary>
        Public Const HangUpPollingInterval As Integer = 50

        ''' <summary>
        ''' Defines the name of the Advapi32.dll.
        ''' </summary>
        Public Const AdvApi32Dll As String = "advapi32.dll"

        ''' <summary>
        ''' Defines the name of the Gdi32.dll.
        ''' </summary>
        Public Const Gdi32Dll As String = "gdi32.dll"

        ''' <summary>
        ''' Defines the name of the Credui.dll.
        ''' </summary>
        Public Const CredUIDll As String = "credui.dll"

        ''' <summary>
        ''' Defines the name of the Rasapi32.dll.
        ''' </summary>
        Public Const RasApi32Dll As String = "rasapi32.dll"

        ''' <summary>
        ''' Defines the name of the Kernel32.dll.
        ''' </summary>
        Public Const Kernel32Dll As String = "kernel32.dll"

        ''' <summary>
        ''' Defines the name of the RasDlg.dll.
        ''' </summary>
        Public Const RasDlgDll As String = "rasdlg.dll"

#Region "Lmcons.h Constants"

        ''' <summary>
        ''' Defines the maximum length of the NetBIOS name.
        ''' </summary>
        Public Const NETBIOS_NAME_LEN As Integer = 16

        ''' <summary>
        ''' Defines the maximum length of a username.
        ''' </summary>
        Public Const UNLEN As Integer = 256

        ''' <summary>
        ''' Defines the maximum length of a password.
        ''' </summary>
        Public Const PWLEN As Integer = 256

        ''' <summary>
        ''' Defines the maximum length of a domain name.
        ''' </summary>
        Public Const DNLEN As Integer = 15

#End Region

#Region "Ras.h Constants"

#Region "RASDT"

        ''' <summary>
        ''' A modem accessed through a COM port.
        ''' </summary>
        Public Const RASDT_Modem As String = "modem"

        ''' <summary>
        ''' An ISDN card with a corresponding NDISWAN driver installed.
        ''' </summary>
        Public Const RASDT_Isdn As String = "isdn"

        ''' <summary>
        ''' An X.25 card with a corresponding NDISWAN driver installed.
        ''' </summary>
        Public Const RASDT_X25 As String = "x25"

        ''' <summary>
        ''' A virtual private network connection.
        ''' </summary>
        Public Const RASDT_Vpn As String = "vpn"

        ''' <summary>
        ''' A packet assembler/disassembler.
        ''' </summary>
        Public Const RASDT_Pad As String = "pad"

        ''' <summary>
        ''' Generic device type.
        ''' </summary>
        Public Const RASDT_Generic As String = "GENERIC"

        ''' <summary>
        ''' Direct serial connection through a serial port.
        ''' </summary>
        Public Const RASDT_Serial As String = "SERIAL"

        ''' <summary>
        ''' Frame Relay.
        ''' </summary>
        Public Const RASDT_FrameRelay As String = "FRAMERELAY"

        ''' <summary>
        ''' Asynchronous Transfer Mode (ATM).
        ''' </summary>
        Public Const RASDT_Atm As String = "ATM"

        ''' <summary>
        ''' Sonet device type.
        ''' </summary>
        Public Const RASDT_Sonet As String = "SONET"

        ''' <summary>
        ''' Switched 56K access.
        ''' </summary>
        Public Const RASDT_SW56 As String = "SW56"

        ''' <summary>
        ''' An Infrared Data Association (IrDA) compliant device.
        ''' </summary>
        Public Const RASDT_Irda As String = "IRDA"

        ''' <summary>
        ''' Direct parallel connection through a parallel port.
        ''' </summary>
        Public Const RASDT_Parallel As String = "PARALLEL"

        ''' <summary>
        ''' Point-to-Point Protocol over Ethernet.
        ''' </summary>
        Public Const RASDT_PPPoE As String = "PPPoE"

#End Region

        ''' <summary>
        ''' Defines the maximum length of a device type.
        ''' </summary>
        Public Const RAS_MaxDeviceType As Integer = 16

        ''' <summary>
        ''' Defines the maximum length of a phone number.
        ''' </summary>
        Public Const RAS_MaxPhoneNumber As Integer = 128

        ''' <summary>
        ''' Defines the maximum length of an IP address.
        ''' </summary>
        Public Const RAS_MaxIpAddress As Integer = 15

        ''' <summary>
        ''' Defines the maximum length of an IPX address.
        ''' </summary>
        Public Const RAS_MaxIpxAddress As Integer = 21

        ''' <summary>
        ''' Defines the maximum length of an entry name.
        ''' </summary>
        Public Const RAS_MaxEntryName As Integer = 256

        ''' <summary>
        ''' Defines the maximum length of a device name.
        ''' </summary>
        Public Const RAS_MaxDeviceName As Integer = 128

        ''' <summary>
        ''' Defines the maximum length of a callback number.
        ''' </summary>
        Public Const RAS_MaxCallbackNumber As Integer = RAS_MaxPhoneNumber

        ''' <summary>
        ''' Defines the maximum length of an area code.
        ''' </summary>
        Public Const RAS_MaxAreaCode As Integer = 10

        ''' <summary>
        ''' Defines the maximum length of a pad type.
        ''' </summary>
        Public Const RAS_MaxPadType As Integer = 32

        ''' <summary>
        ''' Defines the maximum length of an X25 address.
        ''' </summary>
        Public Const RAS_MaxX25Address As Integer = 200

        ''' <summary>
        ''' Defines the maximum length of a facilities.
        ''' </summary>
        Public Const RAS_MaxFacilities As Integer = 200

        ''' <summary>
        ''' Defines the maximum length of a user data.
        ''' </summary>
        Public Const RAS_MaxUserData As Integer = 200

        ''' <summary>
        ''' Defines the maximum length of a reply message.
        ''' </summary>
        Public Const RAS_MaxReplyMessage As Integer = 1024

        ''' <summary>
        ''' Defines the maximum length of a DNS suffix.
        ''' </summary>
        Public Const RAS_MaxDnsSuffix As Integer = 256

        ''' <summary>
        ''' Defines the paused state for a RAS connection state.
        ''' </summary>
        Public Const RASCS_PAUSED As Integer = &H1000

        ''' <summary>
        ''' Defines the done state for a RAS connection state.
        ''' </summary>
        Public Const RASCS_DONE As Integer = &H2000

        ''' <summary>
        ''' Defines the done state for a RAS connection sub-state.
        ''' </summary>
        Public Const RASCSS_DONE As Integer = &H2000

#End Region

#Region "WinCred.h Constants"

#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then

		''' <summary>
		''' Defines the maximum length of the caption for the credential user interface.
		''' </summary>
		Public Const CREDUI_MAX_CAPTION_LENGTH As Integer = 128

		''' <summary>
		''' Defines the maximum length of the message for the credential user interface.
		''' </summary>
		Public Const CREDUI_MAX_MESSAGE_LENGTH As Integer = 32767

		''' <summary>
		''' Defines the maximum length of a username for the credential user interface.
		''' </summary>
		Public Const CREDUI_MAX_USERNAME_LENGTH As Integer = 256 + 1 + 256

		''' <summary>
		''' Defines the maximum length of a username for the credential user interface.
		''' </summary>
		Public Const CREDUI_MAX_PASSWORD_LENGTH As Integer = 512 / 2

#End If

#End Region

#Region "Error Codes"

        ''' <summary>
        ''' Defines the maximum length of a path.
        ''' </summary>
        Public Const MAX_PATH As Integer = 260

        ''' <summary>
        ''' The operation was successful.
        ''' </summary>
        Public Const SUCCESS As Integer = 0

        ''' <summary>
        ''' The system cannot find the file specified.
        ''' </summary>
        Public Const ERROR_FILE_NOT_FOUND As Integer = 2

        ''' <summary>
        ''' The user did not have appropriate permissions to perform the requested action.
        ''' </summary>
        Public Const ERROR_ACCESS_DENIED As Integer = 5

        ''' <summary>
        ''' The handle is invalid.
        ''' </summary>
        Public Const ERROR_INVALID_HANDLE As Integer = 6

        ''' <summary>
        ''' The parameter is incorrect.
        ''' </summary>
        Public Const ERROR_INVALID_PARAMETER As Integer = 87

        '''//// <summary>
        '''//// The filename, directory name, or volume label syntax is incorrect.
        '''//// </summary>
        '''/public const int ERROR_INVALID_NAME = 123;

        ''' <summary>
        ''' Cannot create a file when that file already exists.
        ''' </summary>
        Public Const ERROR_ALREADY_EXISTS As Integer = 183

        ''' <summary>
        ''' The operation was cancelled by the user.
        ''' </summary>
        Public Const ERROR_CANCELLED As Integer = 1223

        '''//// <summary>
        '''//// An operation is pending.
        '''//// </summary>
        '''/public const int PENDING = (RASBASE + 0);

        '''//// <summary>
        '''//// An invalid port handle was detected.
        '''//// </summary>
        '''/public const int ERROR_INVALID_PORT_HANDLE = (RASBASE + 1);

        '''//// <summary>
        '''//// The specified port is already open.
        '''//// </summary>
        '''/public const int ERROR_PORT_ALREADY_OPEN = (RASBASE + 2);

        ''' <summary>
        ''' The caller's buffer is too small.
        ''' </summary>
        Public Const ERROR_BUFFER_TOO_SMALL As Integer = RASBASE + 3

        '''//// <summary>
        '''//// Incorrect information was specified.
        '''//// </summary>
        '''/public const int ERROR_WRONG_INFO_SPECIFIED = (RASBASE + 4);

        '''//// <summary>
        '''//// The port information cannot be set.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_SET_PORT_INFO = (RASBASE + 5);

        '''//// <summary>
        '''//// The specified port is not connected.
        '''//// </summary>
        '''/public const int ERROR_PORT_NOT_CONNECTED = (RASBASE + 6);

        '''//// <summary>
        '''//// An invalid event was detected.
        '''//// </summary>
        '''/public const int ERROR_EVENT_INVALID = (RASBASE + 7);

        '''//// <summary>
        '''//// A device was specified that does not exist.
        '''//// </summary>
        '''/public const int ERROR_DEVICE_DOES_NOT_EXIST = (RASBASE + 8);

        '''//// <summary>
        '''//// A device type was specified that does not exist.
        '''//// </summary>
        '''/public const int ERROR_DEVICETYPE_DOES_NOT_EXIST = (RASBASE + 9);

        '''//// <summary>
        '''//// An invalid buffer was specified.
        '''//// </summary>
        '''/public const int ERROR_BUFFER_INVALID = (RASBASE + 10);

        '''//// <summary>
        '''//// A route was specified that is not available.
        '''//// </summary>
        '''/public const int ERROR_ROUTE_NOT_AVAILABLE = (RASBASE + 11);

        '''//// <summary>
        '''//// A route was specified that is not allocated.
        '''//// </summary>
        '''/public const int ERROR_ROUTE_NOT_ALLOCATED = (RASBASE + 12);

        '''//// <summary>
        '''//// An invalid compression was specified.
        '''//// </summary>
        '''/public const int ERROR_INVALID_COMPRESSION_SPECIFIED = (RASBASE + 13);

        '''//// <summary>
        '''//// There were insufficient buffers available.
        '''//// </summary>
        '''/public const int ERROR_OUT_OF_BUFFERS = (RASBASE + 14);

        '''//// <summary>
        '''//// The specified port was not found.
        '''//// </summary>
        '''/public const int ERROR_PORT_NOT_FOUND = (RASBASE + 15);

        '''//// <summary>
        '''//// An asynchronous request is pending.
        '''//// </summary>
        '''/public const int ERROR_ASYNC_REQUEST_PENDING = (RASBASE + 16);

        '''//// <summary>
        '''//// The modem (or other connecting device) is already disconnecting.
        '''//// </summary>
        '''/public const int ERROR_ALREADY_DISCONNECTING = (RASBASE + 17);

        '''//// <summary>
        '''//// The specified port is not open.
        '''//// </summary>
        '''/public const int ERROR_PORT_NOT_OPEN = (RASBASE + 18);

        '''//// <summary>
        '''//// A connection to the remote computer could not be established, so the port used for this connection was closed.
        '''//// </summary>
        '''/public const int ERROR_PORT_DISCONNECTED = (RASBASE + 19);

        '''//// <summary>
        '''//// No endpoints could be determined.
        '''//// </summary>
        '''/public const int ERROR_NO_ENDPOINTS = (RASBASE + 20);

        ''' <summary>
        ''' The system could not open the phone book file.
        ''' </summary>
        Public Const ERROR_CANNOT_OPEN_PHONEBOOK As Integer = RASBASE + 21

        '''//// <summary>
        '''//// The system could not load the phone book file.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_LOAD_PHONEBOOK = (RASBASE + 22);

        ''' <summary>
        ''' The system could not find the phone book entry for this connection.
        ''' </summary>
        Public Const ERROR_CANNOT_FIND_PHONEBOOK_ENTRY As Integer = RASBASE + 23

        '''//// <summary>
        '''//// The system could not update the phone book file.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_WRITE_PHONEBOOK = (RASBASE + 24);

        '''//// <summary>
        '''//// The system found invalid information in the phone book file.
        '''//// </summary>
        '''/public const int ERROR_CORRUPT_PHONEBOOK = (RASBASE + 25);

        '''//// <summary>
        '''//// A string could not be loaded.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_LOAD_STRING = (RASBASE + 26);

        '''//// <summary>
        '''//// A key could not be found.
        '''//// </summary>
        '''/public const int ERROR_KEY_NOT_FOUND = (RASBASE + 27);

        '''//// <summary>
        '''//// The connection was terminated by the remote computer before it could be completed.
        '''//// </summary>
        '''/public const int ERROR_DISCONNECTION = (RASBASE + 28);

        '''//// <summary>
        '''//// The connection was closed by the remote computer.
        '''//// </summary>
        '''/public const int ERROR_REMOTE_DISCONNECTION = (RASBASE + 29);

        '''//// <summary>
        '''//// The modem (or other connecting device) was disconnected due to hardware failure.
        '''//// </summary>
        '''/public const int ERROR_HARDWARE_FAILURE = (RASBASE + 30);

        '''//// <summary>
        '''//// The user disconnected the modem (or other connecting device).
        '''//// </summary>
        '''/public const int ERROR_USER_DISCONNECTION = (RASBASE + 31);

        '''//// <summary>
        '''//// An incorrect structure size was detected.
        '''//// </summary>
        '''/public const int ERROR_INVALID_SIZE = (RASBASE + 32);

        '''//// <summary>
        '''//// The modem (or other connecting device) is already in use or is not configured properly.
        '''//// </summary>
        '''/public const int ERROR_PORT_NOT_AVAILABLE = (RASBASE + 33);

        '''//// <summary>
        '''//// Your computer could not be registered on the remote network.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_PROJECT_CLIENT = (RASBASE + 34);

        '''//// <summary>
        '''//// There was an unknown error.
        '''//// </summary>
        '''/public const int ERROR_UNKNOWN = (RASBASE + 35);

        '''//// <summary>
        '''//// The device attached to the port is not the one expected.
        '''//// </summary>
        '''/public const int ERROR_WRONG_DEVICE_ATTACHED = (RASBASE + 36);

        '''//// <summary>
        '''//// A string was detected that could not be converted.
        '''//// </summary>
        '''/public const int ERROR_BAD_STRING = (RASBASE + 37);

        '''//// <summary>
        '''//// The remote server is not responding in a timely fashion.
        '''//// </summary>
        '''/public const int ERROR_REQUEST_TIMEOUT = (RASBASE + 38);

        '''//// <summary>
        '''//// No asynchronous net is available.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_GET_LANA = (RASBASE + 39);

        '''//// <summary>
        '''//// An error has occurred involving NetBIOS.
        '''//// </summary>
        '''/public const int ERROR_NETBIOS_ERROR = (RASBASE + 40);

        '''//// <summary>
        '''//// The server cannot allocate NetBIOS resources needed to support the client.
        '''//// </summary>
        '''/public const int ERROR_SERVER_OUT_OF_RESOURCES = (RASBASE + 41);

        '''//// <summary>
        '''//// One of your computer's NetBIOS names is already registered on the remote network.
        '''//// </summary>
        '''/public const int ERROR_NAME_EXISTS_ON_NET = (RASBASE + 42);

        '''//// <summary>
        '''//// A network adapter at the server failed.
        '''//// </summary>
        '''/public const int ERROR_SERVER_GENERAL_NET_FAILURE = (RASBASE + 43);

        '''//// <summary>
        '''//// You will not receive network message popups.
        '''//// </summary>
        '''/public const int WARNING_MSG_ALIAS_NOT_ADDED = (RASBASE + 44);

        '''//// <summary>
        '''//// There was an internal authentication error.
        '''//// </summary>
        '''/public const int ERROR_AUTH_INTERNAL = (RASBASE + 45);

        '''//// <summary>
        '''//// The account is not permitted to log on at this time of day.
        '''//// </summary>
        '''/public const int ERROR_RESTRICTED_LOGON_HOURS = (RASBASE + 46);

        '''//// <summary>
        '''//// The account is disabled.
        '''//// </summary>
        '''/public const int ERROR_ACCT_DISABLED = (RASBASE + 47);

        '''//// <summary>
        '''//// The password for this account has expired.
        '''//// </summary>
        '''/public const int ERROR_PASSWD_EXPIRED = (RASBASE + 48);

        '''//// <summary>
        '''//// The account does not have permission to dial in.
        '''//// </summary>
        '''/public const int ERROR_NO_DIALIN_PERMISSION = (RASBASE + 49);

        '''//// <summary>
        '''//// The remote access server is not responding.
        '''//// </summary>
        '''/public const int ERROR_SERVER_NOT_RESPONDING = (RASBASE + 50);

        '''//// <summary>
        '''//// The modem (or other connecting device) has reported an error.
        '''//// </summary>
        '''/public const int ERROR_FROM_DEVICE = (RASBASE + 51);

        '''//// <summary>
        '''//// There was an unrecognized response from the modem (or other connecting device).
        '''//// </summary>
        '''/public const int ERROR_UNRECOGNIZED_RESPONSE = (RASBASE + 52);

        '''//// <summary>
        '''//// A macro required by the modem (or other connecting device) was not found in the device.INF file.
        '''//// </summary>
        '''/public const int ERROR_MACRO_NOT_FOUND = (RASBASE + 53);

        '''//// <summary>
        '''//// A command or response in the device.INF file section refers to an undefined macro.
        '''//// </summary>
        '''/public const int ERROR_MACRO_NOT_DEFINED = (RASBASE + 54);

        '''//// <summary>
        '''//// The message macro was not found in the device.INF file section.
        '''//// </summary>
        '''/public const int ERROR_MESSAGE_MACRO_NOT_FOUND = (RASBASE + 55);

        '''//// <summary>
        '''//// The defaultoff macro in the device.INF file section contains an undefined macro.
        '''//// </summary>
        '''/public const int ERROR_DEFAULTOFF_MACRO_NOT_FOUND = (RASBASE + 56);

        '''//// <summary>
        '''//// The device.INF file could not be opened.
        '''//// </summary>
        '''/public const int ERROR_FILE_COULD_NOT_BE_OPENED = (RASBASE + 57);

        '''//// <summary>
        '''//// The device name in the device.INF or media.INI file is too long.
        '''//// </summary>
        '''/public const int ERROR_DEVICENAME_TOO_LONG = (RASBASE + 58);

        '''//// <summary>
        '''//// The media.INI file refers to an unknown device name.
        '''//// </summary>
        '''/public const int ERROR_DEVICENAME_NOT_FOUND = (RASBASE + 59);

        '''//// <summary>
        '''//// The device.INF file contains no responses for the command.
        '''//// </summary>
        '''/public const int ERROR_NO_RESPONSES = (RASBASE + 60);

        '''//// <summary>
        '''//// The device.INF file is missing a command.
        '''//// </summary>
        '''/public const int ERROR_NO_COMMAND_FOUND = (RASBASE + 61);

        '''//// <summary>
        '''//// There was an attempt to set a macro not listed in device.INF file section.
        '''//// </summary>
        '''/public const int ERROR_WRONG_KEY_SPECIFIED = (RASBASE + 62);

        '''//// <summary>
        '''//// The media.INI file refers to an unknown device type.
        '''//// </summary>
        '''/public const int ERROR_UNKNOWN_DEVICE_TYPE = (RASBASE + 63);

        '''//// <summary>
        '''//// The system has run out of memory.
        '''//// </summary>
        '''/public const int ERROR_ALLOCATING_MEMORY = (RASBASE + 64);

        '''//// <summary>
        '''//// The modem (or other connecting device) is not properly configured.
        '''//// </summary>
        '''/public const int ERROR_PORT_NOT_CONFIGURED = (RASBASE + 65);

        '''//// <summary>
        '''//// The modem (or other connecting device) is not functioning.
        '''//// </summary>
        '''/public const int ERROR_DEVICE_NOT_READY = (RASBASE + 66);

        '''//// <summary>
        '''//// The system was unable to read the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_INI_FILE = (RASBASE + 67);

        ''' <summary>
        ''' The connection was terminated.
        ''' </summary>
        Public Const ERROR_NO_CONNECTION As Integer = RASBASE + 68

        '''//// <summary>
        '''//// The usage parameter in the media.INI file is invalid.
        '''//// </summary>
        '''/public const int ERROR_BAD_USAGE_IN_INI_FILE = (RASBASE + 69);

        '''//// <summary>
        '''//// The system was unable to read the section name from the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_SECTIONNAME = (RASBASE + 70);

        '''//// <summary>
        '''//// The system was unable to read the device type from the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_DEVICETYPE = (RASBASE + 71);

        '''//// <summary>
        '''//// The system was unable to read the device name from the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_DEVICENAME = (RASBASE + 72);

        '''//// <summary>
        '''//// The system was unable to read the usage from the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_USAGE = (RASBASE + 73);

        '''//// <summary>
        '''//// The system was unable to read the maximum connection BPS rate from the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_MAXCONNECTBPS = (RASBASE + 74);

        '''//// <summary>
        '''//// The system was unable to read the maximum carrier connection speed from the media.INI file.
        '''//// </summary>
        '''/public const int ERROR_READING_MAXCARRIERBPS = (RASBASE + 75);

        '''//// <summary>
        '''//// The phone line is busy.
        '''//// </summary>
        '''/public const int ERROR_LINE_BUSY = (RASBASE + 76);

        '''//// <summary>
        '''//// A person answered instead of a modem (or other connecting device).
        '''//// </summary>
        '''/public const int ERROR_VOICE_ANSWER = (RASBASE + 77);

        '''//// <summary>
        '''//// The remote computer did not respond.
        '''//// </summary>
        '''/public const int ERROR_NO_ANSWER = (RASBASE + 78);

        '''//// <summary>
        '''//// The system could not detect the carrier.
        '''//// </summary>
        '''/public const int ERROR_NO_CARRIER = (RASBASE + 79);

        '''//// <summary>
        '''//// There was no dial tone.
        '''//// </summary>
        '''/public const int ERROR_NO_DIALTONE = (RASBASE + 80);

        '''//// <summary>
        '''//// The modem (or other connecting device) reported a general error.
        '''//// </summary>
        '''/public const int ERROR_IN_COMMAND = (RASBASE + 81);

        '''//// <summary>
        '''//// There was an error in writing the section name.
        '''//// </summary>
        '''/public const int ERROR_WRITING_SECTIONNAME = (RASBASE + 82);

        '''//// <summary>
        '''//// There was an error in writing the device type.
        '''//// </summary>
        '''/public const int ERROR_WRITING_DEVICETYPE = (RASBASE + 83);

        '''//// <summary>
        '''//// There was an error in writing the device name.
        '''//// </summary>
        '''/public const int ERROR_WRITING_DEVICENAME = (RASBASE + 84);

        '''//// <summary>
        '''//// There was an error in writing the maximum connection speed.
        '''//// </summary>
        '''/public const int ERROR_WRITING_MAXCONNECTBPS = (RASBASE + 85);

        '''//// <summary>
        '''//// There was an error in writing the maximum carrier speed.
        '''//// </summary>
        '''/public const int ERROR_WRITING_MAXCARRIERBPS = (RASBASE + 86);

        '''//// <summary>
        '''//// There was an error in writing the usage.
        '''//// </summary>
        '''/public const int ERROR_WRITING_USAGE = (RASBASE + 87);

        '''//// <summary>
        '''//// There was an error in writing the default-off.
        '''//// </summary>
        '''/public const int ERROR_WRITING_DEFAULTOFF = (RASBASE + 88);

        '''//// <summary>
        '''//// There was an error in reading the default-off.
        '''//// </summary>
        '''/public const int ERROR_READING_DEFAULTOFF = (RASBASE + 89);

        '''//// <summary>
        '''//// The INI file was empty.
        '''//// </summary>
        '''/public const int ERROR_EMPTY_INI_FILE = (RASBASE + 90);

        '''//// <summary>
        '''//// Access was denied because the username and/or password was invalid on the domain.
        '''//// </summary>
        '''/public const int ERROR_AUTHENTICATION_FAILURE = (RASBASE + 91);

        '''//// <summary>
        '''//// There was a hardware failure in the modem (or other connecting device).
        '''//// </summary>
        '''/public const int ERROR_PORT_OR_DEVICE = (RASBASE + 92);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_NOT_BINARY_MACRO = (RASBASE + 93);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_DCB_NOT_FOUND = (RASBASE + 94);

        '''//// <summary>
        '''//// The state machines are not started.
        '''//// </summary>
        '''/public const int ERROR_STATE_MACHINES_NOT_STARTED = (RASBASE + 95);

        '''//// <summary>
        '''//// The state machines are already started.
        '''//// </summary>
        '''/public const int ERROR_STATE_MACHINES_ALREADY_STARTED = (RASBASE + 96);

        '''//// <summary>
        '''//// The response looping did not complete.
        '''//// </summary>
        '''/public const int ERROR_PARTIAL_RESPONSE_LOOPING = (RASBASE + 97);

        '''//// <summary>
        '''//// A response keyname in the device.INF file is not in the expected format.
        '''//// </summary>
        '''/public const int ERROR_UNKNOWN_RESPONSE_KEY = (RASBASE + 98);

        '''//// <summary>
        '''//// The modem (or other connecting device) response caused a buffer overflow.
        '''//// </summary>
        '''/public const int ERROR_RECV_BUF_FULL = (RASBASE + 99);

        '''//// <summary>
        '''//// The expanded command in the device.INF file is too long.
        '''//// </summary>
        '''/public const int ERROR_CMD_TOO_LONG = (RASBASE + 100);

        '''//// <summary>
        '''//// The modem moved to a connection speed not supported by the COM driver.
        '''//// </summary>
        '''/public const int ERROR_UNSUPPORTED_BPS = (RASBASE + 101);

        '''//// <summary>
        '''//// Device response received when none expected.
        '''//// </summary>
        '''/public const int ERROR_UNEXPECTED_RESPONSE = (RASBASE + 102);

        ''' <summary>
        ''' The connection needs information from you, but the application does not allow user interaction.
        ''' </summary>
        Public Const ERROR_INTERACTIVE_MODE As Integer = RASBASE + 103

        '''//// <summary>
        '''//// The callback number is invalid.
        '''//// </summary>
        '''/public const int ERROR_BAD_CALLBACK_NUMBER = (RASBASE + 104);

        '''//// <summary>
        '''//// The authorization state is invalid.
        '''//// </summary>
        '''/public const int ERROR_INVALID_AUTH_STATE = (RASBASE + 105);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_WRITING_INITBPS = (RASBASE + 106);

        '''//// <summary>
        '''//// There was an error related to the X.25 protocol.
        '''//// </summary>
        '''/public const int ERROR_X25_DIAGNOSTIC = (RASBASE + 107);

        '''//// <summary>
        '''//// The account has expired.
        '''//// </summary>
        '''/public const int ERROR_ACCT_EXPIRED = (RASBASE + 108);

        '''//// <summary>
        '''//// There was an error changing the password on the domain. The password might have been too short or might 
        '''//// have matched a previously used password.
        '''//// </summary>
        '''/public const int ERROR_CHANGING_PASSWORD = (RASBASE + 109);

        '''//// <summary>
        '''//// Serial overrun errors were detected while communicating with the modem.
        '''//// </summary>
        '''/public const int ERROR_OVERRUN = (RASBASE + 110);

        '''//// <summary>
        '''//// A configuration error on this computer is preventing this connection.
        '''//// </summary>
        '''/public const int ERROR_RASMAN_CANNOT_INITIALIZE = (RASBASE + 111);

        '''//// <summary>
        '''//// The two-way port is initializing. Wait a few seconds and redial.
        '''//// </summary>
        '''/public const int ERROR_BIPLEX_PORT_NOT_AVAILABLE = (RASBASE + 112);

        '''//// <summary>
        '''//// No active ISDN lines are available.
        '''//// </summary>
        '''/public const int ERROR_NO_ACTIVE_ISDN_LINES = (RASBASE + 113);

        '''//// <summary>
        '''//// No ISDN channels are available to make the call.
        '''//// </summary>
        '''/public const int ERROR_NO_ISDN_CHANNELS_AVAILABLE = (RASBASE + 114);

        '''//// <summary>
        '''//// Too many errors occurred because of poor phone line quality.
        '''//// </summary>
        '''/public const int ERROR_TOO_MANY_LINE_ERRORS = (RASBASE + 115);

        '''//// <summary>
        '''//// The remote access service IP configuration is unusable.
        '''//// </summary>
        '''/public const int ERROR_IP_CONFIGURATION = (RASBASE + 116);

        '''//// <summary>
        '''//// No IP addresses are available in the static pool of remote access service IP addresses.
        '''//// </summary>
        '''/public const int ERROR_NO_IP_ADDRESSES = (RASBASE + 117);

        '''//// <summary>
        '''//// The connection was terminated because the remote computer did not respond in a timely manner.
        '''//// </summary>
        '''/public const int ERROR_PPP_TIMEOUT = (RASBASE + 118);

        '''//// <summary>
        '''//// The connection was terminated by the remote computer.
        '''//// </summary>
        '''/public const int ERROR_PPP_REMOTE_TERMINATED = (RASBASE + 119);

        '''//// <summary>
        '''//// A connection to the remote computer could not be established. You might need to change the network 
        '''//// settings for this connection.
        '''//// </summary>
        '''/public const int ERROR_PPP_NO_PROTOCOLS_CONFIGURED = (RASBASE + 120);

        '''//// <summary>
        '''//// The remote computer did not respond.
        '''//// </summary>
        '''/public const int ERROR_PPP_NO_RESPONSE = (RASBASE + 121);

        '''//// <summary>
        '''//// Invalid data was received from the remote computer. This data was ignored.
        '''//// </summary>
        '''/public const int ERROR_PPP_INVALID_PACKET = (RASBASE + 122);

        '''//// <summary>
        '''//// The phone number, including prefix and suffix, is too long.
        '''//// </summary>
        '''/public const int ERROR_PHONE_NUMBER_TOO_LONG = (RASBASE + 123);

        '''//// <summary>
        '''//// The IPX protocol cannot dial out on the modem (or other connecting device) because this computer is 
        '''//// not configured for dialing out (it is an IPX router).
        '''//// </summary>
        '''/public const int ERROR_IPXCP_NO_DIALOUT_CONFIGURED = (RASBASE + 124);

        '''//// <summary>
        '''//// The IPX protocol cannot dial in on the modem (or other connecting device) because this computer is 
        '''//// not configured for dialing in (the IPX router is not installed).
        '''//// </summary>
        '''/public const int ERROR_IPXCP_NO_DIALIN_CONFIGURED = (RASBASE + 125);

        '''//// <summary>
        '''//// The IPX protocol cannot be used for dialing out on more than one modem (or other connecting device) at a time.
        '''//// </summary>
        '''/public const int ERROR_IPXCP_DIALOUT_ALREADY_ACTIVE = (RASBASE + 126);

        '''//// <summary>
        '''//// Cannot access TCPCFG.DLL.
        '''//// </summary>
        '''/public const int ERROR_ACCESSING_TCPCFGDLL = (RASBASE + 127);

        '''//// <summary>
        '''//// The system cannot find an IP adapter.
        '''//// </summary>
        '''/public const int ERROR_NO_IP_RAS_ADAPTER = (RASBASE + 128);

        '''//// <summary>
        '''//// SLIP cannot be used unless the IP protocol is installed.
        '''//// </summary>
        '''/public const int ERROR_SLIP_REQUIRES_IP = (RASBASE + 129);

        '''//// <summary>
        '''//// Computer registration is not complete.
        '''//// </summary>
        '''/public const int ERROR_PROJECTION_NOT_COMPLETE = (RASBASE + 130);

        ''' <summary>
        ''' The protocol is not configured.
        ''' </summary>
        Public Const ERROR_PROTOCOL_NOT_CONFIGURED As Integer = RASBASE + 131

        '''//// <summary>
        '''//// Your computer and the remote computer could not agree on PPP control protocols.
        '''//// </summary>
        '''/public const int ERROR_PPP_NOT_CONVERGING = (RASBASE + 132);

        '''//// <summary>
        '''//// A connection to the remote computer could not be completed. You might need to adjust the protocols on 
        '''//// this computer.
        '''//// </summary>
        '''/public const int ERROR_PPP_CP_REJECTED = (RASBASE + 133);

        '''//// <summary>
        '''//// The PPP link control protocol was terminated.
        '''//// </summary>
        '''/public const int ERROR_PPP_LCP_TERMINATED = (RASBASE + 134);

        '''//// <summary>
        '''//// The requested address was rejected by the server.
        '''//// </summary>
        '''/public const int ERROR_PPP_REQUIRED_ADDRESS_REJECTED = (RASBASE + 135);

        '''//// <summary>
        '''//// The remote computer terminated the control protocol.
        '''//// </summary>
        '''/public const int ERROR_PPP_NCP_TERMINATED = (RASBASE + 136);

        '''//// <summary>
        '''//// Loopback was detected.
        '''//// </summary>
        '''/public const int ERROR_PPP_LOOPBACK_DETECTED = (RASBASE + 137);

        '''//// <summary>
        '''//// The server did not assign an address.
        '''//// </summary>
        '''/public const int ERROR_PPP_NO_ADDRESS_ASSIGNED = (RASBASE + 138);

        '''//// <summary>
        '''//// The authentication protocol required by the remote server cannot use the stored password.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_USE_LOGON_CREDENTIALS = (RASBASE + 139);

        '''//// <summary>
        '''//// An invalid dialing rule was detected.
        '''//// </summary>
        '''/public const int ERROR_TAPI_CONFIGURATION = (RASBASE + 140);

        '''//// <summary>
        '''//// The local computer does not support the required data encryption type.
        '''//// </summary>
        '''/public const int ERROR_NO_LOCAL_ENCRYPTION = (RASBASE + 141);

        '''//// <summary>
        '''//// The remote computer does not support the required data encryption type.
        '''//// </summary>
        '''/public const int ERROR_REMOTE_ENCRYPTION = (RASBASE + 142);

        '''//// <summary>
        '''//// The remote computer requires data encryption.
        '''//// </summary>
        '''/public const int ERROR_REMOTE_REQUIRES_ENCRYPTION = (RASBASE + 143);

        '''//// <summary>
        '''//// The system cannot use the IPX network number assigned by the remote computer.
        '''//// </summary>
        '''/public const int ERROR_IPXCP_NET_NUMBER_CONFLICT = (RASBASE + 144);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_INVALID_SMM = (RASBASE + 145);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_SMM_UNINITIALIZED = (RASBASE + 146);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_NO_MAC_FOR_PORT = (RASBASE + 147);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_SMM_TIMEOUT = (RASBASE + 148);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_BAD_PHONE_NUMBER = (RASBASE + 149);

        '''//// <summary>
        '''//// Undocumented value.
        '''//// </summary>
        '''/public const int ERROR_WRONG_MODULE = (RASBASE + 150);

        '''//// <summary>
        '''//// The callback number contains an invalid character.
        '''//// </summary>
        '''/public const int ERROR_INVALID_CALLBACK_NUMBER = (RASBASE + 151);

        '''//// <summary>
        '''//// A syntax error was encountered while processing a script.
        '''//// </summary>
        '''/public const int ERROR_SCRIPT_SYNTAX = (RASBASE + 152);

        '''//// <summary>
        '''//// The connection could not be disconnected because it was created by the multi-protocol router.
        '''//// </summary>
        '''/public const int ERROR_HANGUP_FAILED = (RASBASE + 153);

        '''//// <summary>
        '''//// The system could not find the multi-link bundle.
        '''//// </summary>
        '''/public const int ERROR_BUNDLE_NOT_FOUND = (RASBASE + 154);

        '''//// <summary>
        '''//// The system cannot perform automated dial because this connection has a custom dialer specified.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_DO_CUSTOMDIAL = (RASBASE + 155);

        '''//// <summary>
        '''//// This connection is already being dialed.
        '''//// </summary>
        '''/public const int ERROR_DIAL_ALREADY_IN_PROGRESS = (RASBASE + 156);

        '''//// <summary>
        '''//// Remote Access Services could not be started automatically.
        '''//// </summary>
        '''/public const int ERROR_RASAUTO_CANNOT_INITIALIZE = (RASBASE + 157);

        '''//// <summary>
        '''//// Internet Connection Sharing is already enabled on the connection.
        '''//// </summary>
        '''/public const int ERROR_CONNECTION_ALREADY_SHARED = (RASBASE + 158);

        '''//// <summary>
        '''//// An error occurred while the existing Internet Connection Sharing settings were being changed.
        '''//// </summary>
        '''/public const int ERROR_SHARING_CHANGE_FAILED = (RASBASE + 159);

        '''//// <summary>
        '''//// An error occurred while routing capabilities were being enabled.
        '''//// </summary>
        '''/public const int ERROR_SHARING_ROUTER_INSTALL = (RASBASE + 160);

        '''//// <summary>
        '''//// An error occurred while Internet Connection Sharing was being enabled for the connection.
        '''//// </summary>
        '''/public const int ERROR_SHARE_CONNECTION_FAILED = (RASBASE + 161);

        '''//// <summary>
        '''//// An error occurred while the local network was being configured for sharing.
        '''//// </summary>
        '''/public const int ERROR_SHARING_PRIVATE_INSTALL = (RASBASE + 162);

        '''//// <summary>
        '''//// Internet Connection Sharing cannot be enabled.  There is more than one LAN connection other than the 
        '''//// connection to be shared.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_SHARE_CONNECTION = (RASBASE + 163);

        '''//// <summary>
        '''//// No smart card reader is installed.
        '''//// </summary>
        '''/public const int ERROR_NO_SMART_CARD_READER = (RASBASE + 164);

        '''//// <summary>
        '''//// Internet Connection Sharing cannot be enabled. A LAN connection is already configured with the IP address 
        '''//// that is required for automatic IP addressing.
        '''//// </summary>
        '''/public const int ERROR_SHARING_ADDRESS_EXISTS = (RASBASE + 165);

        '''//// <summary>
        '''//// A certificate could not be found. Connections that use the L2TP protocol over IPSec require the installation 
        '''//// of a machine certificate, also known as a computer certificate.
        '''//// </summary>
        '''/public const int ERROR_NO_CERTIFICATE = (RASBASE + 166);

        '''//// <summary>
        '''//// Internet Connection Sharing cannot be enabled. The LAN connection selected as the private network has more 
        '''//// than one IP address configured.  Please reconfigure the LAN connection with a single IP address before 
        '''//// enabling Internet Connection Sharing.
        '''//// </summary>
        '''/public const int ERROR_SHARING_MULTIPLE_ADDRESSES = (RASBASE + 167);

        '''//// <summary>
        '''//// The connection attempt failed because of failure to encrypt data.
        '''//// </summary>
        '''/public const int ERROR_FAILED_TO_ENCRYPT = (RASBASE + 168);

        '''//// <summary>
        '''//// The specified destination is not reachable.
        '''//// </summary>
        '''/public const int ERROR_BAD_ADDRESS_SPECIFIED = (RASBASE + 169);

        '''//// <summary>
        '''//// The remote computer rejected the connection attempt.
        '''//// </summary>
        '''/public const int ERROR_CONNECTION_REJECT = (RASBASE + 170);

        '''//// <summary>
        '''//// The connection attempt failed because the network is busy.
        '''//// </summary>
        '''/public const int ERROR_CONGESTION = (RASBASE + 171);

        '''//// <summary>
        '''//// The remote computer's network hardware is incompatible with the type of call requested.
        '''//// </summary>
        '''/public const int ERROR_INCOMPATIBLE = (RASBASE + 172);

        '''//// <summary>
        '''//// The connection attempt failed because the destination number has changed.
        '''//// </summary>
        '''/public const int ERROR_NUMBERCHANGED = (RASBASE + 173);

        '''//// <summary>
        '''//// The connection attempt failed because of a temporary failure.
        '''//// </summary>
        '''/public const int ERROR_TEMPFAILURE = (RASBASE + 174);

        '''//// <summary>
        '''//// The call was blocked by the remote computer.
        '''//// </summary>
        '''/public const int ERROR_BLOCKED = (RASBASE + 175);

        '''//// <summary>
        '''//// The call could not be connected because the remote computer has invoked the Do Not Disturb feature.
        '''//// </summary>
        '''/public const int ERROR_DONOTDISTURB = (RASBASE + 176);

        '''//// <summary>
        '''//// The connection attempt failed because the modem (or other connecting device) on the remote computer is out of order.
        '''//// </summary>
        '''/public const int ERROR_OUTOFORDER = (RASBASE + 177);

        '''//// <summary>
        '''//// It was not possible to verify the identity of the server.
        '''//// </summary>
        '''/public const int ERROR_UNABLE_TO_AUTHENTICATE_SERVER = (RASBASE + 178);

        '''//// <summary>
        '''//// To dial out using this connection you must use a smart card.
        '''//// </summary>
        '''/public const int ERROR_SMART_CARD_REQUIRED = (RASBASE + 179);

        ''' <summary>
        ''' An attempted function is not valid for this connection.
        ''' </summary>
        Public Const ERROR_INVALID_FUNCTION_FOR_ENTRY As Integer = RASBASE + 180

        '''//// <summary>
        '''//// The connection requires a certificate, and no valid certificate was found.
        '''//// </summary>
        '''/public const int ERROR_CERT_FOR_ENCRYPTION_NOT_FOUND = (RASBASE + 181);

        '''//// <summary>
        '''//// Internet Connection Sharing (ICS) and Internet Connection Firewall (ICF) cannot be enabled because
        '''//// Routing and Remote Access has been enabled on this computer.
        '''//// </summary>
        '''/public const int ERROR_SHARING_RRAS_CONFLICT = (RASBASE + 182);

        '''//// <summary>
        '''//// Internet Connection Sharing cannot be enabled. The LAN connection selected as the private network is 
        '''//// either not present, or is disconnected from the network.
        '''//// </summary>
        '''/public const int ERROR_SHARING_NO_PRIVATE_LAN = (RASBASE + 183);

        '''//// <summary>
        '''//// You cannot dial using this connection at logon time, because it is configured to use a user name different 
        '''//// than the one on the smart card.
        '''//// </summary>
        '''/public const int ERROR_NO_DIFF_USER_AT_LOGON = (RASBASE + 184);

        '''//// <summary>
        '''//// You cannot dial using this connection at logon time, because it is not configured to use a smart card.
        '''//// </summary>
        '''/public const int ERROR_NO_REG_CERT_AT_LOGON = (RASBASE + 185);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because there is no valid machine certificate on your computer for 
        '''//// security authentication.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_NO_CERT = (RASBASE + 186);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because the security layer could not authenticate the remote computer.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_AUTH_FAIL = (RASBASE + 187);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because the security layer could not negotiate compatible parameters 
        '''//// with the remote computer.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_ATTRIB_FAIL = (RASBASE + 188);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because the security layer encountered a processing error during initial 
        '''//// negotiations with the remote computer.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_GENERAL_PROCESSING = (RASBASE + 189);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because certificate validation on the remote computer failed.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_NO_PEER_CERT = (RASBASE + 190);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because security policy for the connection was not found.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_NO_POLICY = (RASBASE + 191);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because security negotiation timed out.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_TIMED_OUT = (RASBASE + 192);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because an error occurred while negotiating security.
        '''//// </summary>
        '''/public const int ERROR_OAKLEY_ERROR = (RASBASE + 193);

        '''//// <summary>
        '''//// The Framed Protocol RADIUS attribute for this user is not PPP.
        '''//// </summary>
        '''/public const int ERROR_UNKNOWN_FRAMED_PROTOCOL = (RASBASE + 194);

        '''//// <summary>
        '''//// The Tunnel Type RADIUS attribute for this user is not correct.
        '''//// </summary>
        '''/public const int ERROR_WRONG_TUNNEL_TYPE = (RASBASE + 195);

        '''//// <summary>
        '''//// The Service Type RADIUS attribute for this user is neither Framed nor Callback Framed.
        '''//// </summary>
        '''/public const int ERROR_UNKNOWN_SERVICE_TYPE = (RASBASE + 196);

        '''//// <summary>
        '''//// A connection to the remote computer could not be established because the modem was not found or was busy.
        '''//// </summary>
        '''/public const int ERROR_CONNECTING_DEVICE_NOT_FOUND = (RASBASE + 197);

        '''//// <summary>
        '''//// A certificate could not be found that can be used with this Extensible Authentication Protocol.
        '''//// </summary>
        '''/public const int ERROR_NO_EAPTLS_CERTIFICATE = (RASBASE + 198);

        '''//// <summary>
        '''//// Internet Connection Sharing (ICS) cannot be enabled due to an IP address conflict on the network.
        '''//// </summary>
        '''/public const int ERROR_SHARING_HOST_ADDRESS_CONFLICT = (RASBASE + 199);

        '''//// <summary>
        '''//// Unable to establish the VPN connection. The VPN server may be unreachable, or security parameters may not 
        '''//// be configured properly for this connection.
        '''//// </summary>
        '''/public const int ERROR_AUTOMATIC_VPN_FAILED = (RASBASE + 200);

        '''//// <summary>
        '''//// This connection is configured to validate the identity of the access server, but Windows cannot verify the digital 
        '''//// certificate sent by the server.
        '''//// </summary>
        '''/public const int ERROR_VALIDATING_SERVER_CERT = (RASBASE + 201);

        '''//// <summary>
        '''//// The card supplied was not recognized. Please check that the card is inserted correctly, and fits tightly.
        '''//// </summary>
        '''/public const int ERROR_READING_SCARD = (RASBASE + 202);

        '''//// <summary>
        '''//// The PEAP configuration stored in the session cookie does not match the current session configuration.
        '''//// </summary>
        '''/public const int ERROR_INVALID_PEAP_COOKIE_CONFIG = (RASBASE + 203);

        '''//// <summary>
        '''//// The PEAP identity stored in the session cookie does not match the current identity.
        '''//// </summary>
        '''/public const int ERROR_INVALID_PEAP_COOKIE_USER = (RASBASE + 204);

        '''//// <summary>
        '''//// You cannot dial using this connection at logon time, because it is configured to use logged on user's credentials.
        '''//// </summary>
        '''/public const int ERROR_INVALID_MSCHAPV2_CONFIG = (RASBASE + 205);

        '''//// <summary>
        '''//// The VPN connection between your computer and the VPN server could not be completed.
        '''//// </summary>
        '''/public const int ERROR_VPN_GRE_BLOCKED = (RASBASE + 206);

        '''//// <summary>
        '''//// The network connection between your computer and the VPN server was interrupted.
        '''//// </summary>
        '''/public const int ERROR_VPN_DISCONNECT = (RASBASE + 207);

        '''//// <summary>
        '''//// The network connection between your computer and the VPN server could not be established because the remote server refused the connection.
        '''//// </summary>
        '''/public const int ERROR_VPN_REFUSED = (RASBASE + 208);

        '''//// <summary>
        '''//// The network connection between your computer and the VPN server could not be established because the remote server is not responding.
        '''//// </summary>
        '''/public const int ERROR_VPN_TIMEOUT = (RASBASE + 209);

        '''//// <summary>
        '''//// A network connection between your computer and the VPN server was started, but the VPN connection was not completed.
        '''//// </summary>
        '''/public const int ERROR_VPN_BAD_CERT = (RASBASE + 210);

        '''//// <summary>
        '''//// The network conection between your computer and the VPN server could not be established because the remote server is not responding.
        '''//// </summary>
        '''/public const int ERROR_VPN_BAD_PSK = (RASBASE + 211);

        '''//// <summary>
        '''//// The connection was prevented because of a policy configured on your RAS/VPN server.
        '''//// </summary>
        '''/public const int ERROR_SERVER_POLICY = (RASBASE + 212);

        '''//// <summary>
        '''//// You have attempted to establish a second broadband connection while a previous broadband connection is already established using the same device or port.
        '''//// </summary>
        '''/public const int ERROR_BROADBAND_ACTIVE = (RASBASE + 213);

        '''//// <summary>
        '''//// The underlying Ethernet connectivity required for the broadband connection was not found.
        '''//// </summary>
        '''/public const int ERROR_BROADBAND_NO_NIC = (RASBASE + 214);

        '''//// <summary>
        '''//// The broadband network conection could not be established on your computer because the remote server is not responding.
        '''//// </summary>
        '''/public const int ERROR_BROADBAND_TIMEOUT = (RASBASE + 215);

        '''//// <summary>
        '''//// A feature or setting you have tried to enable is no longer supported by the remote access service.
        '''//// </summary>
        '''/public const int ERROR_FEATURE_DEPRECATED = (RASBASE + 216);

        '''//// <summary>
        '''//// Cannot delete a connection while it is connected.
        '''//// </summary>
        '''/public const int ERROR_CANNOT_DELETE = (RASBASE + 217);

        '''//// <summary>
        '''//// The Network Access Protection (NAP) enforcement client could not create system resources for remote access connections.
        '''//// </summary>
        '''/public const int ERROR_RASQEC_RESOURCE_CREATION_FAILED = (RASBASE + 218);

        '''//// <summary>
        '''//// The Network Access Protection Agent (NAPAgent) service has been disabled or is not installed on this computer.
        '''//// </summary>
        '''/public const int ERROR_RASQEC_NAPAGENT_NOT_ENABLED = (RASBASE + 219);

        '''//// <summary>
        '''//// The Network Access Protection (NAP) enforcement client failed to register with the Network Access Protection Agent (NAPAgent) service.
        '''//// </summary>
        '''/public const int ERROR_RASQEC_NAPAGENT_NOT_CONNECTED = (RASBASE + 220);

        '''//// <summary>
        '''//// The Network Access Protection (NAP) enforcement client was unable to process the request because the remote access connection does not exist.
        '''//// </summary>
        '''/public const int ERROR_RASQEC_CONN_DOESNOTEXIST = (RASBASE + 221);

        '''//// <summary>
        '''//// The Network Access Protection (NAP) enforcement client did not respond.
        '''//// </summary>
        '''/public const int ERROR_RASQEC_TIMEOUT = (RASBASE + 222);

        '''//// <summary>
        '''//// Received Crypto-Binding TLV is invalid.
        '''//// </summary>
        '''/public const int ERROR_PEAP_CRYPTOBINDING_INVALID = (RASBASE + 223);

        '''//// <summary>
        '''//// Crypto-Binding TLV is not received.
        '''//// </summary>
        '''/public const int ERROR_PEAP_CRYPTOBINDING_NOTRECEIVED = (RASBASE + 224);

        '''//// <summary>
        '''//// Point-to-Point Tunneling Protocol (PPTP) is incompatible with IPv6.
        '''//// </summary>
        '''/public const int ERROR_INVALID_VPNSTRATEGY = (RASBASE + 225);

        '''//// <summary>
        '''//// EAPTLS validation of the cached credentials failed.
        '''//// </summary>
        '''/public const int ERROR_EAPTLS_CACHE_CREDENTIALS_INVALID = (RASBASE + 226);

        '''//// <summary>
        '''//// The L2TP/IPsec connection cannot be completed because the IKE and AuthIP IPSec Keying Modules service and/or the Base Filtering Engine service is not running.
        '''//// </summary>
        '''/public const int ERROR_IPSEC_SERVICE_STOPPED = (RASBASE + 227);

        '''//// <summary>
        '''//// The connection was terminated because of idle timeout.
        '''//// </summary>
        '''/public const int ERROR_IDLE_TIMEOUT = (RASBASE + 228);

        '''//// <summary>
        '''//// The modem (or other connecting device) was disconnected due to link failure.
        '''//// </summary>
        '''/public const int ERROR_LINK_FAILURE = (RASBASE + 229);

        '''//// <summary>
        '''//// The connection was terminated because user logged off.
        '''//// </summary>
        '''/public const int ERROR_USER_LOGOFF = (RASBASE + 230);

        '''//// <summary>
        '''//// The connection was terminated because user switch happened.
        '''//// </summary>
        '''/public const int ERROR_FAST_USER_SWITCH = (RASBASE + 231);

        '''//// <summary>
        '''//// The connection was terminated because of hibernation.
        '''//// </summary>
        '''/public const int ERROR_HIBERNATION = (RASBASE + 232);

        '''//// <summary>
        '''//// The connection was terminated because the system got suspended.
        '''//// </summary>
        '''/public const int ERROR_SYSTEM_SUSPENDED = (RASBASE + 233);

        '''//// <summary>
        '''//// The connection was terminated because Remote Access Connection manager stopped.
        '''//// </summary>
        '''/public const int ERROR_RASMAN_SERVICE_STOPPED = (RASBASE + 234);

        '''//// <summary>
        '''//// The L2TP connection attempt failed because the security layer could not authenticate the remote computer.
        '''//// </summary>
        '''/public const int ERROR_INVALID_SERVER_CERT = (RASBASE + 235);

        ''' <summary>
        ''' The connection is not configured for network access protection.
        ''' </summary>
        Public Const ERROR_NOT_NAP_CAPABLE As Integer = RASBASE + 236

#End Region

        ''' <summary>
        ''' Defines the invalid handle value.
        ''' </summary>
        '   Public Shared ReadOnly INVALID_HANDLE_VALUE As New RasHandle(New IntPtr(-1), False)

        ''' <summary>
        ''' Defines the base for all RAS error codes.
        ''' </summary>
        Private Const RASBASE As Integer = 600

#End Region

#Region "Delegates"

        ''' <summary>
        ''' The callback used by the RasDial function when a change of state occurs during a remote access connection process.
        ''' </summary>
        ''' <param name="callbackId">An application defined value that was passed to the remote access service.</param>
        ''' <param name="subEntryId">The one-based subentry index for the phone book entry associated with this connection.</param>
        ''' <param name="handle">The handle to the connection.</param>
        ''' <param name="message">The type of event that has occurred.</param>
        ''' <param name="state">The state the remote access connection process is about to enter.</param>
        ''' <param name="errorCode">The error that has occurred. If no error has occurred the value is zero.</param>
        ''' <param name="extendedErrorCode">Any extended error information for certain non-zero values of <paramref name="errorCode"/>.</param>
        ''' <returns><b>true</b> to continue to receive callback notifications, otherwise <b>false</b>.</returns>
        ' Public Delegate Function RasDialFunc2(ByVal callbackId As IntPtr, ByVal subEntryId As Integer, ByVal handle As IntPtr, ByVal message As Integer, ByVal state As RasConnectionState, ByVal errorCode As Integer, _
        ' ByVal extendedErrorCode As Integer) As Boolean

        ''' <summary>
        ''' The callback used by the RasPhonebookDlg function that receives notifications of user activity while the dialog box is open.
        ''' </summary>
        ''' <param name="callbackId">An application defined value that was passed to the RasPhonebookDlg function.</param>
        ''' <param name="eventType">The event that occurred.</param>
        ''' <param name="text">A string whose value depends on the <paramref name="eventType"/> parameter.</param>
        ''' <param name="data">Pointer to an additional buffer argument whose value depends on the <paramref name="eventType"/> parameter.</param>
        Public Delegate Sub RasPBDlgFunc(ByVal callbackId As IntPtr, ByVal eventType As RASPBDEVENT, ByVal text As String, ByVal data As IntPtr)

#End Region

#Region "Enums"

#Region "CREDUI_FLAGS"

        ''' <summary>
        ''' Defines the credential user interface flags.
        ''' </summary>
        <Flags()> _
        Public Enum CREDUI_FLAGS
            ''' <summary>
            ''' No flags have been specified.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Notify the user of insufficient credentials by displaying the 'Logon unsuccessful' balloon tip.
            ''' </summary>
            IncorrectPassword = &H1

            ''' <summary>
            ''' Do not store credentials or display check boxes. You can pass CREDUI_FLAGS_SHOW_SAVE_CHECK_BOX with this flag to display the Save check box only, and the result is returned in the pfSave output parameter.
            ''' </summary>
            DoNotPersist = &H2

            ''' <summary>
            ''' Populate the combo box with local administrators only.
            ''' </summary>
            RequestAdministrator = &H4

            ''' <summary>
            ''' Populate the combo box with user name/password only. Do not display certificates or smart cards in the combo box.
            ''' </summary>
            ExcludeCertificates = &H8

            ''' <summary>
            ''' Populate the combo box with certificates and smart cards only. Do not allow a user name to be entered.
            ''' </summary>
            RequireCertificate = &H10

            ''' <summary>
            ''' If the check box is selected, show the Save check box and return TRUE in the pfSave output parameter, otherwise, return FALSE. Check box uses the value in pfSave by default.
            ''' </summary>
            ShowSaveCheckBox = &H40

            ''' <summary>
            ''' Specifies that a user interface will be shown even if the credentials can be returned from an existing credential in credential manager. This flag is permitted only if CREDUI_FLAGS_GENERIC_CREDENTIALS is also specified.
            ''' </summary>
            AlwaysShowUI = &H80

            ''' <summary>
            ''' Populate the combo box with certificates or smart cards only. Do not allow a user name to be entered.
            ''' </summary>
            RequireSmartCard = &H100

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            PasswordOnlyOK = &H200

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            ValidateUserName = &H400

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            CompleteUserName = &H800

            ''' <summary>
            ''' Do not show the Save check box, but the credential is saved as though the box were shown and selected.
            ''' </summary>
            Persist = &H1000

            ''' <summary>
            ''' When this flag is specified, wildcard credentials will not be matched.
            ''' </summary>
            ServerCredential = &H4000

            ''' <summary>
            ''' Do not persist unless caller later confirms credential via the CredUIConfirmCredential API.
            ''' </summary>
            ExpectConfirmation = &H20000

            ''' <summary>
            ''' Consider the credentials entered by the user to be generic credentials.
            ''' </summary>
            GenericCredentials = &H40000

            ''' <summary>
            ''' The credential is a "run-as" credential. The TargetName parameter specifies the name of the command or program being run. It is used for prompting purposes only.
            ''' </summary>
            UserNameTargetCredentials = &H80000

            ''' <summary>
            ''' Do not allow the user to change the supplied username.
            ''' </summary>
            KeepUserName = &H100000
        End Enum

#End Region

#Region "RASADP"

        ''' <summary>
        ''' Defines the remote access service (RAS) AutoDial parameters.
        ''' </summary>
        Public Enum RASADP
            ''' <summary>
            ''' Causes AutoDial to disable a dialog box displayed to the user before dialing a connection.
            ''' </summary>
            DisableConnectionQuery = 0

            ''' <summary>
            ''' Causes the system to disable all AutoDial connections for the current logon session.
            ''' </summary>
            LogOnSessionDisable

            ''' <summary>
            ''' Indicates the maximum number of addresses that AutoDial stores in the registry.
            ''' </summary>
            SavedAddressesLimit

            ''' <summary>
            ''' Indicates a timeout value (in seconds) when an AutoDial connection attempt fails before another connection attempt begins.
            ''' </summary>
            FailedConnectionTimeout

            ''' <summary>
            ''' Indicates a timeout value (in seconds) when the system displays a dialog box asking the user to confirm that the system should dial.
            ''' </summary>
            ConnectionQueryTimeout
        End Enum

#End Region

#Region "RASAPIVERSION"

        ''' <summary>
        ''' Defines the version of a data structure.
        ''' </summary>
        Public Enum RASAPIVERSION
            ''' <summary>
            ''' The Windows 2000 version.
            ''' </summary>
            Win2K = 1

            ''' <summary>
            ''' The Windows XP version.
            ''' </summary>
            WinXP = 2

            ''' <summary>
            ''' The Windows Vista version (includes Vista SP1). 
            ''' </summary>
            WinVista = 3

            ''' <summary>
            ''' The Windows 7 version.
            ''' </summary>
            Win7 = 4
        End Enum

#End Region

#Region "RASCCPO"

        ''' <summary>
        ''' Defines the remote access service compression control protocol options.
        ''' </summary>
        Public Enum RASCCPO
            ''' <summary>
            ''' No compression options in use.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Compression without encryption.
            ''' </summary>
            CompressionOnly = &H1

            ''' <summary>
            ''' Microsoft Point-to-Point Encryption (MPPE) in stateless mode.
            ''' </summary>
            ''' <remarks>The session key is changed after every packet. This mode improves performance on high latency networks, or networks that experience significant packet loss.</remarks>
            HistoryLess = &H2

            ''' <summary>
            ''' Microsoft Point-to-Point Encryption (MPPE) using 56 bit keys.
            ''' </summary>
            Encryption56Bit = &H10

            ''' <summary>
            ''' Microsoft Point-to-Point Encryption (MPPE) using 40 bit keys.
            ''' </summary>
            Encryption40Bit = &H20

            ''' <summary>
            ''' Microsoft Point-to-Point Encryption (MPPE) using 128 bit keys.
            ''' </summary>
            Encryption128Bit = &H40
        End Enum

#End Region

#Region "RASCF"

#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then

		''' <summary>
		''' Defines the remote access connection options.
		''' </summary>
		<Flags> _
		Public Enum RASCF
			''' <summary>
			''' No connection options specified.
			''' </summary>
			None = &H0

			''' <summary>
			''' Specifies the connection is available to all users.
			''' </summary>
			AllUsers = &H1

			''' <summary>
			''' Specifies the credentials used for the connection are the default credentials.
			''' </summary>
			GlobalCredentials = &H2

			''' <summary>
			''' Specifies the owner of the connection is known.
			''' </summary>
			OwnerKnown = &H4

			''' <summary>
			''' Specifies the owner of the connection matches the current user.
			''' </summary>
			OwnerMatch = &H8
		End Enum

#End If

#End Region

#Region "RASCM"

        ''' <summary>
        ''' Defines the flags indicating which members of a <see cref="RASCREDENTIALS"/> instance are valid.
        ''' </summary>
        <Flags()> _
        Public Enum RASCM
            ''' <summary>
            ''' No options are valid.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' The user name member is valid.
            ''' </summary>
            UserName = &H1

            ''' <summary>
            ''' The password member is valid.
            ''' </summary>
            Password = &H2

            ''' <summary>
            ''' The domain name member is valid.
            ''' </summary>
            Domain = &H4
#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then
			''' <summary>
			''' Indicates the credentials are the default credentials for an all-user connection.
			''' </summary>
			DefaultCredentials = &H8

			''' <summary>
			''' Indicates a pre-shared key should be retrieved.
			''' </summary>
			PreSharedKey = &H10

			''' <summary>
			''' Used to set the pre-shared key on the remote access server.
			''' </summary>
			ServerPreSharedKey = &H20

			''' <summary>
			''' Used to set the pre-shared key for a demand dial interface.
			''' </summary>
			DdmPreSharedKey = &H40
#End If
        End Enum

#End Region

#Region "RASCN"

        ''' <summary>
        ''' Defines the RAS connection event types.
        ''' </summary>
        <Flags()> _
        Public Enum RASCN
            ''' <summary>
            ''' Signal when a connection is created.
            ''' </summary>
            Connection = &H1

            ''' <summary>
            ''' Signal when a connection has disconnected.
            ''' </summary>
            Disconnection = &H2

            ''' <summary>
            ''' Signal when a connection has bandwidth added.
            ''' </summary>
            BandwidthAdded = &H4

            ''' <summary>
            ''' Signal when a connection has bandwidth removed. 
            ''' </summary>
            BandwidthRemoved = &H8

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            Dormant = &H10

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            Reconnection = &H20
        End Enum

#End Region

#Region "RASEAPF"

        ''' <summary>
        ''' Defines flags that are used to qualify the authentication process.
        ''' </summary>
        <Flags()> _
        Public Enum RASEAPF
            ''' <summary>
            ''' No flags are used.
            ''' </summary>
            None = 0

            ''' <summary>
            ''' Specifies that the authentication protocol should not bring up a graphical user interface. If this flag is not present,
            ''' it is okay for the protocol to display a user interface.
            ''' </summary>
            NonInteractive = &H2

            ''' <summary>
            ''' Specifies that the user data is obtained from WinLogon.
            ''' </summary>
            LogOn = &H4

            ''' <summary>
            ''' Specifies that the user should be prompted for identity information before dialing.
            ''' </summary>
            Preview = &H8
        End Enum

#End Region

#Region "RASEO"

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

#End Region

#Region "RASEO2"

#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then

		''' <summary>
		''' Defines the additional connection options for entries.
		''' </summary>
		<Flags> _
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
			#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
			''' <summary>
			''' Indicates the routing compartments feature is enabled.
			''' </summary>
			SecureRoutingCompartment = &H400

			''' <summary>
			''' Configures a VPN connection to use the typical settings for authentication and encryption for the RAS connection.
			''' </summary>
			UseTypicalSettings = &H800

			''' <summary>
			''' Uses the IPv6 DNS address and alternate DNS address for the connection.
			''' </summary>
			IPv6SpecificNameServer = &H1000

			''' <summary>
			''' Indicates the default route for IPv6 packets is through the PPP connection when the connection is active.
			''' </summary>
			IPv6RemoteDefaultGateway = &H2000

			''' <summary>
			''' Registers the IP address with the DNS server when connected.
			''' </summary>
			RegisterIPWithDns = &H4000

			''' <summary>
			''' The DNS suffix for this connection should be used for DNS registration.
			''' </summary>
			UseDnsSuffixForRegistration = &H8000

			''' <summary>
			''' Indicates the administrator is allowed to statically set the interface metric of the IPv4 stack for this interface.
			''' </summary>
			IPv4ExplicitMetric = &H10000

			''' <summary>
			''' Indicates the administrator is allowed to statically set the interface metric of the IPv6 stack for this interface.
			''' </summary>
			IPv6ExplicitMetric = &H20000

			''' <summary>
			''' The IKE validation check will not be performed.
			''' </summary>
			DisableIkeNameEkuCheck = &H40000
			#End If
			#If (WIN7 OrElse WIN8) Then
			''' <summary>
			''' Indicates whether a class based route based on the VPN interface IP address will not be added.
			''' </summary>
			DisableClassBasedStaticRoute = &H80000

			''' <summary>
			''' Indicates whether the specific IPv6 address will be used for the connection.
			''' </summary>
			''' <remarks>If set, RAS attempts to use the IP address specified by <see cref="RasEntry.IPv6Address"/> as the IPv6 address for the connection.</remarks>
			IPv6SpecificAddress = &H100000

			''' <summary>
			''' Indicates whether the client will not be able to change the external IP address of the IKEv2 VPN connection.
			''' </summary>
			DisableMobility = &H200000

			''' <summary>
			''' Indicates whether machine certificates are used for IKEv2 authentication.
			''' </summary>
			RequireMachineCertificates = &H400000
			#End If
			#If (WIN8) Then
			''' <summary>
			''' Use the pre-shared key for the IKEv2 initiator.
			''' </summary>
			UsePreSharedKeyForIkev2Initiator = &H800000

			''' <summary>
			''' Use the pre-shared key for the IKEv2 responder.
			''' </summary>
			UsePreSharedKeyForIkev2Responder = &H1000000

			''' <summary>
			''' Indicates the credentials should be cached.
			''' </summary>
			CacheCredentials = &H2000000
			#End If
		End Enum

#End If

#End Region

#Region "RASIPO"

        ''' <summary>
        ''' Defines the remote access service (RAS) IPCP options.
        ''' </summary>
        <Flags()> _
        Public Enum RASIPO
            ''' <summary>
            ''' No options in use.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Van Jacobson compression.
            ''' </summary>
            VJ = &H1
        End Enum

#End Region

#Region "RASLCPO"

        ''' <summary>
        ''' Defines the Link Control Protocol (LCP) options.
        ''' </summary>
        <Flags()> _
        Public Enum RASLCPO
            ''' <summary>
            ''' No LCP options used.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Protocol Field Compression.
            ''' </summary>
            Pfc = &H1

            ''' <summary>
            ''' Address and Control Field Compression.
            ''' </summary>
            Acfc = &H2

            ''' <summary>
            ''' Short Sequence Number Header Format.
            ''' </summary>
            Sshf = &H4

            ''' <summary>
            ''' DES 56-bit encryption.
            ''' </summary>
            Des56 = &H8

            ''' <summary>
            ''' Triple DES encryption.
            ''' </summary>
            TripleDes = &H10

#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then

			''' <summary>
			''' AES 128-bit encryption.
			''' </summary>
			Aes128 = &H20

			''' <summary>
			''' AES 256-bit encryption.
			''' </summary>
			Aes256 = &H40

#End If
        End Enum

#End Region

#Region "RASNP"

        ''' <summary>
        ''' Defines the network protocols used for negotiation.
        ''' </summary>
        <Flags()> _
        Public Enum RASNP
            ''' <summary>
            ''' No network protocol specified.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Negotiate the NetBEUI protocol.
            ''' </summary>
            NetBeui = &H1

            ''' <summary>
            ''' Negotiate the IPX protocol.
            ''' </summary>
            Ipx = &H2

            ''' <summary>
            ''' Negotiate the IPv4 protocol.
            ''' </summary>
            IP = &H4

#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then

			''' <summary>
			''' Negotiate the IPv6 protocol.
			''' </summary>
			IPv6 = &H8

#End If
        End Enum

#End Region

#Region "RASPBDEVENT"

        ''' <summary>
        ''' Defines the events that occur for RasPhoneBookDialog boxes.
        ''' </summary>
        Public Enum RASPBDEVENT
            ''' <summary>
            ''' Received when the user creates a new phone book entry or copies an existing phone book entry.
            ''' </summary>
            AddEntry = 1

            ''' <summary>
            ''' Received when the user changes an existing phone book entry.
            ''' </summary>
            EditEntry = 2

            ''' <summary>
            ''' Received when the user deletes a phone book entry.
            ''' </summary>
            RemoveEntry = 3

            ''' <summary>
            ''' Received when the user successfully dials an entry.
            ''' </summary>
            DialEntry = 4

            ''' <summary>
            ''' Received when the user makes changes in the <b>User Preferences</b> property sheet.
            ''' </summary>
            EditGlobals = 5

            ''' <summary>
            ''' Received during the dialog box initialization when the NoUser flag has been set.
            ''' </summary>
            NoUser = 6

            ''' <summary>
            ''' Received when the NoUser flag has been set and the user changes the credentials that are supplied during the NoUser event.
            ''' </summary>
            NoUserEdit = 7
        End Enum

#End Region

#Region "RASPROJECTION"

        ''' <summary>
        ''' Defines the projection types.
        ''' </summary>
        Public Enum RASPROJECTION
            ''' <summary>
            ''' Authentication Message Block (AMB) protocol.
            ''' </summary>
            Amb = &H10000

            ''' <summary>
            ''' NetBEUI Framer (NBF) protocol.
            ''' </summary>
            Nbf = &H803F

            ''' <summary>
            ''' Internetwork Packet Exchange (IPX) control protocol.
            ''' </summary>
            Ipx = &H802B

            ''' <summary>
            ''' Internet Protocol (IP) control protocol.
            ''' </summary>
            IP = &H8021

            ''' <summary>
            ''' Compression Control Protocol (CCP).
            ''' </summary>
            Ccp = &H80FD

            ''' <summary>
            ''' Link Control Protocol (LCP).
            ''' </summary>
            Lcp = &HC021

            ''' <summary>
            ''' Internet Protocol Version 6 (IPv6) control protocol.
            ''' </summary>
            IPv6 = &H8057

            ''' <summary>
            ''' Serial Line Internet Protocol (SLIP).
            ''' </summary>
            Slip = &H20000
        End Enum

#End Region

#Region "RASPROJECTION_INFO_TYPE"

#If (WIN7 OrElse WIN8) Then
		''' <summary>
		''' Defines the structure type available for projection operations.
		''' </summary>
		Public Enum RASPROJECTION_INFO_TYPE
			''' <summary>
			''' The projection contains a <see cref="NativeMethods.RASPPP_PROJECTION_INFO"/> structure.
			''' </summary>
			Ppp = 1

			''' <summary>
			''' The project contains a <see cref="NativeMethods.RASIKEV2_PROJECTION_INFO"/> structure.
			''' </summary>
			IkeV2 = 2
		End Enum
#End If

#End Region

#Region "RASDDFLAG"

        ''' <summary>
        ''' Defines the options available for dial dialog boxes.
        ''' </summary>
        <Flags()> _
        Public Enum RASDDFLAG As UInteger
            ''' <summary>
            ''' Use the values specified by the top and left members to position the dialog box.
            ''' </summary>
            PositionDlg = &H1

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            NoPrompt = &H2

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            AoacRedial = &H4

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            LinkFailure = &H80000000UI
        End Enum

#End Region

#Region "RASEDFLAG"

        ''' <summary>
        ''' Defines the options available for entry dialog boxes.
        ''' </summary>
        <Flags()> _
        Public Enum RASEDFLAG
            ''' <summary>
            ''' Use the values specified by the top and left members to position the dialog box.
            ''' </summary>
            PositionDlg = &H1

            ''' <summary>
            ''' Display a wizard for creating a new phone book entry.
            ''' </summary>
            NewEntry = &H2
#If (Not WIN2K8 AndAlso Not WIN7 AndAlso Not WIN8) Then
            ''' <summary>
            ''' Create a new entry by copying the properties of an existing entry.
            ''' </summary>
            CloneEntry = &H4
#End If
            ''' <summary>
            ''' Displays a property sheet for editing the properties of a phone book entry without the ability to rename the entry.
            ''' </summary>
            NoRename = &H8

            ''' <summary>
            ''' Reserved for system use.
            ''' </summary>
            ShellOwned = &H40000000

            ''' <summary>
            ''' Causes the wizard to go directly to the page for a phone line entry.
            ''' </summary>
            NewPhoneEntry = &H10

            ''' <summary>
            ''' Causes the wizard to go directly to the page for a new tunnel entry.
            ''' </summary>
            NewTunnelEntry = &H20
#If (Not WIN2K8 AndAlso Not WIN7 AndAlso Not WIN8) Then
            ''' <summary>
            ''' Causes the wizard to go directly to the page for a direct-connection entry.
            ''' </summary>
            NewDirectEntry = &H40
#End If
            ''' <summary>
            ''' Causes the wizard that creates a new broadband phone book entry.
            ''' </summary>
            NewBroadbandEntry = &H80

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            InternetEntry = &H100

            ''' <summary>
            ''' This member is undocumented.
            ''' </summary>
            NAT = &H200
#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
			''' <summary>
			''' Displays a wizard that allows a user to enable and configure incoming connections.
			''' </summary>
			IncomingConnection = &H400
#End If
        End Enum

#End Region

#Region "RASIKEV2"

#If (WIN7 OrElse WIN8) Then

		''' <summary>
		''' Defines the Internet Key Exchange (IKEv2) options.
		''' </summary>
		<Flags> _
		Public Enum RASIKEV2
			''' <summary>
			''' No options specified.
			''' </summary>
			None = 0

			''' <summary>
			''' The client supports the IKEv2 Mobility and Multi-homing (MOBIKE) protocol.
			''' </summary>
			MobileIke = &H1

			''' <summary>
			''' The client is behind Network Address Translation (NAT).
			''' </summary>
			ClientBehindNat = &H2

			''' <summary>
			''' The server is behind Network Address Translation (NAT).
			''' </summary>
			ServerBehindNat = &H4
		End Enum

#End If

#End Region

#Region "RASPBDLGFLAG"

        ''' <summary>
        ''' Defines the options available for the phone book dialog box.
        ''' </summary>
        <Flags()> _
        Public Enum RASPBDFLAG As UInteger
            ''' <summary>
            ''' Use the values specified by the top and left members to position the dialog box.
            ''' </summary>
            PositionDlg = &H1

            ''' <summary>
            ''' Turns on the close-on-dial option, overriding the user's preference.
            ''' </summary>
            ForceCloseOnDial = &H2

            ''' <summary>
            ''' Causes the dialog callback function to receive a NoUser notification when the dialog box is starting up.
            ''' </summary>
            NoUser = &H10

            ''' <summary>
            ''' Causes the default window position to be saved on exit.
            ''' </summary>
            UpdateDefaults = &H80000000UI
        End Enum

#End Region

#Region "RASTUNNELENDPOINTTYPE"

#If (WIN7 OrElse WIN8) Then

		''' <summary>
		''' Defines the endpoint types.
		''' </summary>
		Public Enum RASTUNNELENDPOINTTYPE
			''' <summary>
			''' The endpoint type is unknown.
			''' </summary>
			Unknown

			''' <summary>
			''' The endpoint type is IPv4.
			''' </summary>
			IPv4

			''' <summary>
			''' The endpoint type is IPv6.
			''' </summary>
			IPv6
		End Enum

#End If

#End Region

#Region "RasNotifierType"

        ''' <summary>
        ''' Defines the callback signatures available for the dialing process.
        ''' </summary>
        Public Enum RasNotifierType
            ''' <summary>
            ''' This member is not supported.
            ''' </summary>
            <Obsolete("Use the RasDialFunc2 value.", True)> _
            RasDialFunc = 0

            ''' <summary>
            ''' This member is not supported.
            ''' </summary>
            <Obsolete("Use the RasDialFunc2 value.", True)> _
            RasDialFunc1 = 1

            ''' <summary>
            ''' The callback function is a <see cref="NativeMethods.RasDialFunc2"/> delegate.
            ''' </summary>
            RasDialFunc2 = 2
        End Enum

#End Region

#Region "RDEOPT"

        ''' <summary>
        ''' Defines the remote access dial extensions options.
        ''' </summary>
        <Flags()> _
        Public Enum RDEOPT
            ''' <summary>
            ''' No dial extension options have been specified.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Use the prefix and suffix that is in the RAS phone book.
            ''' </summary>
            ''' <remarks>If no entry name was specified during dialing, this member is ignored.</remarks>
            UsePrefixSuffix = &H1

            ''' <summary>
            ''' Accept paused states.
            ''' </summary>
            ''' <remarks>If this flag has not been set, the dialer reports a fatal error if it enters a paused state.</remarks>
            PausedStates = &H2

            ''' <summary>
            ''' Ignore the modem speaker setting that is in the RAS phone book, and use the setting specified by the <see cref="SetModemSpeaker"/> flag.
            ''' </summary>
            IgnoreModemSpeaker = &H4

            ''' <summary>
            ''' Sets the modem speaker on.
            ''' </summary>
            SetModemSpeaker = &H8

            ''' <summary>
            ''' Ignore the software compression setting that is in the RAS phone book, and uses the setting specified by the <see cref="SetSoftwareCompression"/> flag.
            ''' </summary>
            IgnoreSoftwareCompression = &H10

            ''' <summary>
            ''' Use software compression.
            ''' </summary>
            SetSoftwareCompression = &H20

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            DisableConnectedUI = &H40

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            DisableReconnectUI = &H80

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            DisableReconnect = &H100

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            NoUser = &H200

            ''' <summary>
            ''' Used internally by the <see cref="RasDialDialog"/> class so that a Windows-95-style logon script is executed in a terminal window visible to the user.
            ''' </summary>
            ''' <remarks>Applications should not set this flag.</remarks>
            PauseOnScript = &H400

            ''' <summary>
            ''' Undocumented flag.
            ''' </summary>
            Router = &H800

            ''' <summary>
            ''' Dial normally instead of calling the RasCustomDial entry point of the custom dialer.
            ''' </summary>
            CustomDial = &H1000

#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then

			''' <summary>
			''' Specifies that RasDial should invoke a custom-scripting DLL after establishing the connection to the server.
			''' </summary>
			UseCustomScripting = &H2000

#End If
        End Enum

#End Region

#End Region

#Region "Structures"

#Region "CREDUI_INFO"

#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then

		''' <summary>
		''' Describes a dialog box used to obtain user credentials.
		''' </summary>
		<StructLayout(LayoutKind.Sequential, Pack := 4)> _
		Public Structure CREDUI_INFO
			Public size As Integer
			Public hwndOwner As IntPtr
			<MarshalAs(UnmanagedType.LPTStr)> _
			Public message As String
			<MarshalAs(UnmanagedType.LPTStr)> _
			Public caption As String
			Public bannerBitmap As IntPtr
		End Structure

#End If

#End Region

#Region "RAS_STATS"

        ''' <summary>
        ''' Describes statistics for a remote access service (RAS) connection.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)> _
        Public Structure RAS_STATS
            Public size As Integer
            Public bytesTransmitted As UInteger
            Public bytesReceived As UInteger
            Public framesTransmitted As UInteger
            Public framesReceived As UInteger
            Public crcError As UInteger
            Public timeoutError As UInteger
            Public alignmentError As UInteger
            Public hardwareOverrunError As UInteger
            Public framingError As UInteger
            Public bufferOverrunError As UInteger
            Public compressionRatioIn As UInteger
            Public compressionRatioOut As UInteger
            Public linkSpeed As UInteger
            Public connectionDuration As UInteger
        End Structure

#End Region

#Region "RASAMB"

        ''' <summary>
        ''' Describes the result of a remote access server (RAS) Authentication Message Block (AMB) projection operation.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASAMB
            Public size As Integer
            Public errorCode As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=NETBIOS_NAME_LEN + 1)> _
            Public netBiosError As String
            Public lana As Byte
        End Structure

#End Region

#Region "RASAUTODIALENTRY"

        ''' <summary>
        ''' Describes an AutoDial entry associated with a network address in the AutoDial mapping database.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASAUTODIALENTRY
            Public size As Integer
            Public options As Integer
            Public dialingLocation As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> _
            Public entryName As String
        End Structure

#End Region

#Region "RASCTRYINFO"

        ''' <summary>
        ''' Describes country/region specific dialing information.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Public Structure RASCTRYINFO
            Public size As Integer
            Public countryId As Integer
            Public nextCountryId As Integer
            Public countryCode As Integer
            Public countryNameOffset As Integer
        End Structure

#End Region

#Region "RASEAPUSERIDENTITY"

        ''' <summary>
        ''' Describes extensible authentication protocol (EAP) user identity information for a particular user.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASEAPUSERIDENTITY
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=UNLEN + 1)> _
            Public userName As String
            Public sizeOfEapData As Integer
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)> _
            Public eapData As Byte()
        End Structure

#End Region

#Region "RASEAPINFO"

        ''' <summary>
        ''' Describes user-specific Extensible Authentication Protocol (EAP) information.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, Pack:=4)> _
        Public Structure RASEAPINFO
            Public sizeOfEapData As Integer
            Public eapData As IntPtr
        End Structure

#End Region

#Region "RASNAPSTATE"

#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
		''' <summary>
		''' Describes the network access protection (NAP) state of a connection.
		''' </summary>
		<StructLayout(LayoutKind.Sequential, CharSet := CharSet.Unicode)> _
		Public Structure RASNAPSTATE
			Public size As Integer
			Public flags As Integer
			Public isolationState As RasIsolationState
			Public probationTime As FILETIME
		End Structure
#End If

#End Region

#Region "RASPPPCCP"

        ''' <summary>
        ''' Describes the results of a Compression Control Protocol (CCP) projection operation.
        ''' </summary>
        '<StructLayout(LayoutKind.Sequential)> _
        'Public Structure RASPPPCCP
        '    Public size As Integer
        '    Public errorCode As Integer
        '    Public compressionAlgorithm As RasCompressionType
        '    Public options As NativeMethods.RASCCPO
        '    Public serverCompressionAlgorithm As RasCompressionType
        '    Public serverOptions As NativeMethods.RASCCPO
        'End Structure

#End Region

#Region "RASPPPIP"

        ''' <summary>
        ''' Describes the results of a PPP IP projection operation.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASPPPIP
            Public size As Integer
            Public errorCode As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxIpAddress + 1)> _
            Public ipAddress As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxIpAddress + 1)> _
            Public serverIPAddress As String
            Public options As RASIPO
            Public serverOptions As RASIPO
        End Structure

#End Region

#Region "RASPPPIPX"

        ''' <summary>
        ''' Describes the results of a PPP IPX projection operation.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASPPPIPX
            Public size As Integer
            Public errorCode As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxIpxAddress + 1)> _
            Public ipxAddress As String
        End Structure

#End Region

#Region "RASPPPLCP"

        ''' <summary>
        ''' Describes the results of a PPP Link Control Protocol (LCP) multilink projection operation
        ''' </summary>
        '<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        'Public Structure RASPPPLCP
        '    Public size As Integer
        '    <MarshalAs(UnmanagedType.Bool)> _
        '    Public bundled As Boolean
        '    Public errorCode As Integer
        '    Public authenticationProtocol As RasLcpAuthenticationType
        '    Public authenticationData As RasLcpAuthenticationDataType
        '    Public eapTypeId As Integer
        '    Public serverAuthenticationProtocol As RasLcpAuthenticationType
        '    Public serverAuthenticationData As RasLcpAuthenticationDataType
        '    Public serverEapTypeId As Integer
        '    <MarshalAs(UnmanagedType.Bool)> _
        '    Public multilink As Boolean
        '    Public terminateReason As Integer
        '    Public serverTerminateReason As Integer
        '    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxReplyMessage)> _
        '    Public replyMessage As String
        '    Public options As NativeMethods.RASLCPO
        '    Public serverOptions As NativeMethods.RASLCPO
        'End Structure

#End Region

#Region "RASPPPNBF"

        ''' <summary>
        ''' Describes the result of a PPP NetBEUI Framer (NBF) projection operation.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASPPPNBF
            Public size As Integer
            Public errorCode As Integer
            Public netBiosErrorCode As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=NETBIOS_NAME_LEN + 1)> _
            Public netBiosErrorMessage As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=NETBIOS_NAME_LEN + 1)> _
            Public workstationName As String
            Public lana As Byte
        End Structure

#End Region

#Region "RASSLIP"

        ''' <summary>
        ''' Describes the result of a Serial Line Internet Protocol (SLIP) projection operation.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Public Structure RASSLIP
            Public size As Integer
            Public errorCode As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxIpAddress + 1)> _
            Public ipAddress As String
        End Structure

#End Region

#Region "RASPPPIPV6"

#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
		''' <summary>
		''' Describes the results of a Point-to-Point (PPP) IPv6 projection operation..
		''' </summary>
		<StructLayout(LayoutKind.Sequential, Pack := 4)> _
		Public Structure RASPPPIPV6
			Public size As Integer
			Public errorCode As Integer
			<MarshalAs(UnmanagedType.ByValArray, SizeConst := 8)> _
			Public localInterfaceIdentifier As Byte()
			<MarshalAs(UnmanagedType.ByValArray, SizeConst := 8)> _
			Public peerInterfaceIdentifier As Byte()
			<MarshalAs(UnmanagedType.ByValArray, SizeConst := 2)> _
			Public localCompressionProtocol As Byte()
			<MarshalAs(UnmanagedType.ByValArray, SizeConst := 2)> _
			Public peerCompressionProtocol As Byte()
		End Structure
#End If

#End Region

#Region "RASCONN"

        ''' <summary>
        ''' Describes a remote access connection.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASCONN
            Public size As Integer
            Public handle As IntPtr
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> _
            Public entryName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceType + 1)> _
            Public deviceType As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceName + 1)> _
            Public deviceName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
            Public phoneBook As String
            Public subEntryId As Integer
            Public entryId As Guid
#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then
			Public connectionOptions As NativeMethods.RASCF
			Public sessionId As Luid
#End If
#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
			Public correlationId As Guid
#End If
        End Structure

#End Region

        '#Region "RASCONNSTATUS"

        '        ''' <summary>
        '        ''' Describes the current status of a remote access connection.
        '        ''' </summary>
        '        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        '        Public Structure RASCONNSTATUS
        '            Public size As Integer
        '            Public connectionState As RasConnectionState
        '            Public errorCode As Integer
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceType + 1)> _
        '            Public deviceType As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceName + 1)> _
        '            Public deviceName As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxPhoneNumber + 1)> _
        '            Public phoneNumber As String

        '#If (WIN7 OrElse WIN8) Then
        '			Public localEndPoint As RASTUNNELENDPOINT
        '			Public remoteEndPoint As RASTUNNELENDPOINT
        '			Public connectionSubState As RasConnectionSubState
        '#End If
        '        End Structure

        '#End Region

#Region "RASCREDENTIALS"

        ''' <summary>
        ''' Describes user credentials associated with a phone book entry.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASCREDENTIALS
            Public size As Integer
            Public options As RASCM
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=UNLEN + 1)> _
            Public userName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=PWLEN + 1)> _
            Public password As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=DNLEN + 1)> _
            Public domain As String
        End Structure

#End Region

#Region "RASDEVINFO"

        ''' <summary>
        ''' Describes a TAPI device capable of establishing a remote access service connection.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Public Structure RASDEVINFO
            Public size As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceType + 1)> _
            Public type As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceName + 1)> _
            Public name As String
        End Structure

#End Region

#Region "RASDEVSPECIFICINFO"

        ''' <summary>
        ''' Describes a cookie for server validation and bypass point-to-point (PPP) authentication.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, Pack:=4)> _
        Public Structure RASDEVSPECIFICINFO
            Public size As Integer
            Public cookie As IntPtr
        End Structure

#End Region

#Region "RASDIALDLG"

        ''' <summary>
        ''' Specifies additional input and output parameters for the RasDialDlg function.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)> _
        Public Structure RASDIALDLG
            Public size As Integer
            Public hwndOwner As IntPtr
            Public flags As RASDDFLAG
            Public left As Integer
            Public top As Integer
            Public subEntryId As Integer
            Public [error] As Integer
            Public reserved As IntPtr
            Public reserved2 As IntPtr
        End Structure

#End Region

#Region "RASDIALEXTENSIONS"

        ''' <summary>
        ''' Contains information about extended features of the RasDial function.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, Pack:=4)> _
        Public Structure RASDIALEXTENSIONS
            Public size As Integer
            Public options As NativeMethods.RDEOPT
            Public handle As IntPtr
            Public reserved As IntPtr
            Public reserved1 As IntPtr
            Public eapInfo As RASEAPINFO
#If (WIN7 OrElse WIN8) Then
			Public skipPppAuth As Boolean
			Public devSpecificInfo As RASDEVSPECIFICINFO
#End If
        End Structure

#End Region

#Region "RASDIALPARAMS"

        ''' <summary>
        ''' Contains information used by RasDial to establish a remote access connection.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASDIALPARAMS
            Public size As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> _
            Public entryName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxPhoneNumber + 1)> _
            Public phoneNumber As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxCallbackNumber + 1)> _
            Public callbackNumber As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=UNLEN + 1)> _
            Public userName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=PWLEN + 1)> _
            Public password As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=DNLEN + 1)> _
            Public domain As String
            Public subEntryId As Integer
            Public callbackId As IntPtr
#If (WIN7 OrElse WIN8) Then
			Public interfaceIndex As Integer
#End If
            '''/#if (WIN8)
            '''/            public IntPtr szEncPassword;
            '''/#endif
        End Structure

#End Region

        '#Region "RASENTRY"

        '        ''' <summary>
        '        ''' Describes a phone book entry.
        '        ''' </summary>
        '        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        '        Public Structure RASENTRY
        '            Public size As Integer
        '            Public options As NativeMethods.RASEO

        '            ' Location and phone number
        '            Public countryId As Integer
        '            Public countryCode As Integer
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxAreaCode + 1)> _
        '            Public areaCode As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxPhoneNumber + 1)> _
        '            Public phoneNumber As String
        '            Public alternateOffset As Integer

        '            ' PPP/IP
        '            Public ipAddress As RASIPADDR
        '            Public dnsAddress As RASIPADDR
        '            Public dnsAddressAlt As RASIPADDR
        '            Public winsAddress As RASIPADDR
        '            Public winsAddressAlt As RASIPADDR

        '            ' Framing
        '            Public frameSize As Integer
        '            Public networkProtocols As NativeMethods.RASNP
        '            ' Public framingProtocol As RasFramingProtocol

        '            ' Scripting 
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
        '            Public script As String

        '            ' AutoDial
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
        '            Public autoDialDll As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
        '            Public autoDialFunc As String

        '            ' Device
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceType + 1)> _
        '            Public deviceType As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceName + 1)> _
        '            Public deviceName As String

        '            ' X.25
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxPadType + 1)> _
        '            Public x25PadType As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxX25Address + 1)> _
        '            Public x25Address As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxFacilities + 1)> _
        '            Public x25Facilities As String
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxUserData + 1)> _
        '            Public x25UserData As String
        '            Public channels As Integer

        '            ' Reserved
        '            Public reserved1 As Integer
        '            Public reserved2 As Integer

        '            ' Multilink and BAP
        '            Public subentries As Integer
        '            ' Public dialMode As RasDialMode
        '            Public dialExtraPercent As Integer
        '            Public dialExtraSampleSeconds As Integer
        '            Public hangUpExtraPercent As Integer
        '            Public hangUpExtraSampleSeconds As Integer

        '            ' Idle timeout
        '            Public idleDisconnectSeconds As Integer

        '            '  Public entryType As RasEntryType
        '            'Public encryptionType As RasEncryptionType
        '            Public customAuthKey As Integer
        '            Public id As Guid
        '            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
        '            Public customDialDll As String
        '            ' Public vpnStrategy As RasVpnStrategy

        '#If (WINXP OrElse WIN2K8 OrElse WIN7 OrElse WIN8) Then
        '			Public options2 As NativeMethods.RASEO2
        '			Public options3 As Integer
        '			<MarshalAs(UnmanagedType.ByValTStr, SizeConst := RAS_MaxDnsSuffix)> _
        '			Public dnsSuffix As String
        '			Public tcpWindowSize As Integer
        '			<MarshalAs(UnmanagedType.ByValTStr, SizeConst := MAX_PATH)> _
        '			Public prerequisitePhoneBook As String
        '			<MarshalAs(UnmanagedType.ByValTStr, SizeConst := RAS_MaxEntryName + 1)> _
        '			Public prerequisiteEntryName As String
        '			Public redialCount As Integer
        '			Public redialPause As Integer
        '#End If
        '#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
        '			Public ipv6DnsAddress As RASIPV6ADDR
        '			Public ipv6DnsAddressAlt As RASIPV6ADDR
        '			Public ipv4InterfaceMetric As Integer
        '			Public ipv6InterfaceMetric As Integer
        '#End If
        '#If (WIN7 OrElse WIN8) Then
        '			Public ipv6Address As RASIPV6ADDR
        '			Public ipv6PrefixLength As Integer
        '			Public networkOutageTime As Integer
        '#End If
        '        End Structure

        '#End Region

#Region "RASENTRYDLG"

        ''' <summary>
        ''' Specifies additional input and output parameters for the RasEntryDlg function.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASENTRYDLG
            Public size As Integer
            Public hwndOwner As IntPtr
            Public flags As RASEDFLAG
            Public left As Integer
            Public top As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> _
            Public entryName As String
            Public [error] As Integer
            Public reserved As IntPtr
            Public reserved2 As IntPtr
        End Structure

#End Region

#Region "RASENTRYNAME"

        ''' <summary>
        ''' Describes an entry name from a remote access phone book.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASENTRYNAME
            Public size As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> _
            Public name As String
            ' Public phoneBookType As RasPhoneBookType
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH + 1)> _
            Public phoneBookPath As String
        End Structure

#End Region

#Region "RASIKEV2_PROJECTION_INFO"

#If (WIN7 OrElse WIN8) Then

		''' <summary>
		''' Contains projection information obtained during Internet Key Exchange (IKE) negotiation.
		''' </summary>
		<StructLayout(LayoutKind.Sequential, Pack := 4)> _
		Public Structure RASIKEV2_PROJECTION_INFO
			Public ipv4NegotiationError As Integer
			Public ipv4Address As RASIPADDR
			Public ipv4ServerAddress As RASIPADDR
			Public ipv6NegotiationError As Integer
			Public ipv6Address As RASIPV6ADDR
			Public ipv6ServerAddress As RASIPV6ADDR
			Public prefixLength As Integer
			Public authenticationProtocol As RasIkeV2AuthenticationType
			Public eapTypeId As Integer
			Public options As NativeMethods.RASIKEV2
			Public encryptionMethod As RasIPSecEncryptionType
			Public numIPv4ServerAddresses As Integer
			Public ipv4ServerAddresses As IntPtr
			Public numIPv6ServerAddresses As Integer
			Public ipv6ServerAddresses As IntPtr
		End Structure

#End If

#End Region

#Region "RASIPADDR"

        ''' <summary>
        ''' Describes an IPv4 address.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)> _
        <TypeConverter(GetType(IPAddressConverter))> _
        Public Structure RASIPADDR
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)> _
            Public addr As Byte()
        End Structure

#End Region

#Region "RASIPV6ADDR"

#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then

		''' <summary>
		''' Describes an IPv6 address.
		''' </summary>
		<StructLayout(LayoutKind.Sequential)> _
		<TypeConverter(GetType(IPAddressConverter))> _
		Public Structure RASIPV6ADDR
			<MarshalAs(UnmanagedType.ByValArray, SizeConst := 16)> _
			Public addr As Byte()
		End Structure

#End If

#End Region

#Region "RASPBDLG"

        ''' <summary>
        ''' Specifies additional input and output parameters for the RasPhonebookDlg function.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, Pack:=4)> _
        Public Structure RASPBDLG
            Public size As Integer
            Public hwndOwner As IntPtr
            Public flags As RASPBDFLAG
            Public left As Integer
            Public top As Integer
            Public callbackId As IntPtr
            Public callback As RasPBDlgFunc
            Public [error] As Integer
            Public reserved As IntPtr
            Public reserved2 As IntPtr
        End Structure

#End Region

#Region "RAS_PROJECTION_INFO"

#If (WIN7 OrElse WIN8) Then
		''' <summary>
		''' Contains projection information obtained during Point-to-Point (PPP) negotiation.
		''' </summary>
		''' <remarks>This structure does not match the SDK due to a union on the object and the portability problems that it causes on 64-bit platforms. See the implementation for additional details.</remarks>
		<StructLayout(LayoutKind.Sequential)> _
		Public Structure RAS_PROJECTION_INFO
			Public version As RASAPIVERSION
			Public type As RASPROJECTION_INFO_TYPE
		End Structure
#End If

#End Region

#Region "RASPPP_PROJECTION_INFO"

        ''' <summary>
        ''' Contains information obtained during Point-to-Point (PPP) negotiation.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)> _
        Public Structure RASPPP_PROJECTION_INFO
            Public ipv4NegotiationError As Integer
            Public ipv4Address As RASIPADDR
            Public ipv4ServerAddress As RASIPADDR
            Public ipv4Options As RasIPOptions
            Public ipv4ServerOptions As RasIPOptions
            Public ipv6NegotiationError As Integer
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> _
            Public interfaceIdentifier As Byte()
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> _
            Public serverInterfaceIdentifier As Byte()
            <MarshalAs(UnmanagedType.Bool)> _
            Public bundled As Boolean
            <MarshalAs(UnmanagedType.Bool)> _
            Public multilink As Boolean
            Public authenticationProtocol As RasLcpAuthenticationType
            Public authenticationData As RasLcpAuthenticationDataType
            Public serverAuthenticationProtocol As RasLcpAuthenticationType
            Public serverAuthenticationData As RasLcpAuthenticationDataType
            Public eapTypeId As Integer
            Public serverEapTypeId As Integer
            Public lcpOptions As NativeMethods.RASLCPO
            Public serverLcpOptions As NativeMethods.RASLCPO
            Public ccpCompressionAlgorithm As RasCompressionType
            Public serverCcpCompressionAlgorithm As RasCompressionType
            Public ccpOptions As NativeMethods.RASCCPO
            Public serverCcpOptions As NativeMethods.RASCCPO
        End Structure

#End Region

#Region "RASSUBENTRY"

        ''' <summary>
        ''' Describes a subentry of a remote access service (RAS) phone book entry.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=4)> _
        Public Structure RASSUBENTRY
            Public size As Integer
            Public flags As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceType + 1)> _
            Public deviceType As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceName + 1)> _
            Public deviceName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxPhoneNumber + 1)> _
            Public phoneNumber As String
            Public alternateOffset As Integer
        End Structure

#End Region

#Region "RASTUNNELENDPOINT"

#If (WIN7 OrElse WIN8) Then
		''' <summary>
		''' Describes the endpoint of a virtual private network (VPN) tunnel.
		''' </summary>
		<StructLayout(LayoutKind.Sequential)> _
		Public Structure RASTUNNELENDPOINT
			Public type As RASTUNNELENDPOINTTYPE
			<MarshalAs(UnmanagedType.ByValArray, SizeConst := 16)> _
			Public addr As Byte()
		End Structure
#End If

#End Region

#Region "RASUPDATECONN"

#If (WIN7 OrElse WIN8) Then
		''' <summary>
		''' Contains data used to update an active remote access service (RAS) connection.
		''' </summary>
		<StructLayout(LayoutKind.Sequential)> _
		Public Structure RASUPDATECONN
			Public version As RASAPIVERSION
			Public size As Integer
			Public flags As Integer
			Public interfaceIndex As Integer
			Public localEndPoint As RASTUNNELENDPOINT
			Public remoteEndPoint As RASTUNNELENDPOINT
		End Structure
#End If

#End Region

#End Region
    End Class
End Namespace