Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Text

Imports System

Imports Microsoft.VisualBasic


Public Class ClassMdmInstall

    Public Event finishInstallDeviceDriver(ByVal esito As Boolean)
    Public Event TreeviewPortUpdated(ByVal nPortInstalled As Integer)
    Private txtLog As TextBox = Nothing
    Private _lblReadingRawString As Label = Nothing
    Private _btnRefresh As Button
    Private _tablLytPanelUnistall As TableLayoutPanel
    Private _txtUsbRaw As TextBox = Nothing
    Private PathLog As String = ""
    Private nomeFileLog As String = "LogDriverInstall.txt"

    Public listPortFTDI As New List(Of classPortFTDI)

    Public listMdmInstalledName As New List(Of LibRasBook.ClassRasBook.classMdmInstalled)

    Private Const keyFtdibus As String = "SYSTEM\CurrentControlSet\Enum\FTDIBUS"


    Private _IslistingModem As Boolean = False


    Class classPortFTDI
        Public COM As String = ""
        Public IsPlugged As Boolean = False
        Public DescPid As String = ""
        Public PID As String = ""
        Public Usb As ClassSerachHeminaUSBraw.classUSBraw

        Public Shared Function GetNumPortsPlugged(l As List(Of classPortFTDI)) As Integer
            Try
                If l Is Nothing Then Return 0
                If l.Count = 0 Then Return 0
                Dim n As Integer = 0
                For i As Integer = 0 To l.Count - 1
                    If l(i).IsPlugged = True Then
                        n += 1
                    End If
                Next
                Return n
            Catch ex As Exception

            End Try
        End Function

    End Class


#Region "DllImport"

    'Marshaling with C# – Chapter 2: Marshaling Simple Types
    'http://www.codeproject.com/Articles/66244/Marshaling-with-C-Chapter-Marshaling-Simple-Type

    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiGetDeviceInstallParams(ByVal hDevinfo As IntPtr,
                              ByRef DeviceInfoData As SP_DEVINFO_DATA,
                              ByRef DeviceInstallParams As SP_DEVINSTALL_PARAMS
                              ) As Boolean
    End Function

    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiSetDeviceInstallParams(ByVal hDevinfo As IntPtr,
                              ByRef DeviceInfoData As SP_DEVINFO_DATA,
                              ByRef DeviceInstallParams As SP_DEVINSTALL_PARAMS
                              ) As Boolean
    End Function


    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiBuildDriverInfoList(ByVal hDevinfo As IntPtr,
                              ByRef DeviceInfoData As SP_DEVINFO_DATA,
                              ByVal DriverType As Integer
                              ) As Boolean
    End Function


    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiDestroyDriverInfoList(ByVal hDevinfo As IntPtr,
                              ByRef DeviceInfoData As SP_DEVINFO_DATA,
                              ByVal DriverType As Integer
                              ) As Boolean
    End Function

    <DllImport("setupapi.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function SetupDiEnumDriverInfo(
<[In]> ByVal lpInfoSet As IntPtr,
<[In], [Optional]> ByVal deviceInfoData As SP_DEVINFO_DATA,
<[In]> ByVal driverType As UInt32,
<[In]> ByVal memberIndex As Int32,
<Out> ByRef driverInfoData As SP_DRVINFO_DATA) As Boolean

    End Function

    '
    <DllImport("setupapi.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function SetupDiSetSelectedDriver(
<[In]> ByVal lpInfoSet As IntPtr,
<[In], [Optional]> ByVal deviceInfoData As SP_DEVINFO_DATA,
<Out> ByRef driverInfoData As SP_DRVINFO_DATA) As Boolean

    End Function

    Public Const DIF_SELECTBESTCOMPATDRV As Integer = &H17
    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiCallClassInstaller(ByVal InstallFunction As IntPtr,
                          ByVal DeviceInfoSet As IntPtr,
                          ByRef DeviceInfoData As SP_DEVINFO_DATA
                         ) As Boolean
    End Function

    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiRemoveDevice(ByVal DeviceInfoSet As IntPtr,
                          ByRef DeviceInfoData As SP_DEVINFO_DATA
                         ) As Boolean
    End Function
    <DllImport("setupapi.dll")>
    Public Shared Function SetupDiSetSelectedDevice(ByVal DeviceInfoSet As IntPtr,
                          ByRef DeviceInfoData As SP_DEVINFO_DATA
                         ) As Boolean
    End Function
    Public Const DIF_REGISTERDEVICE As Integer = &H19

    Public Const SPDIT_NODRIVER As Integer = &H0
    Public Const SPDIT_CLASSDRIVER As Integer = &H1
    Public Const SPDIT_COMPATDRIVER As Integer = &H2


    '<StructLayout(LayoutKind.Sequential, Pack:=1)>
    'Public Structure SP_DEVINSTALL_PARAMS
    '    Public cbSize As Integer
    '    Public Flags As Integer
    '    Public FlagsEx As Integer
    '    Public hwndParent As IntPtr
    '    Public InstallMsgHandler As IntPtr
    '    Public InstallMsgHandlerContext As IntPtr
    '    Public FileQueue As IntPtr
    '    Public ClassInstallReserved As IntPtr
    '    Public Reserved As Integer
    '    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX__DESC)>
    '    Public DriverPath As StringBuilder
    'End Structure


    '<StructLayout(LayoutKind.Sequential, Pack:=1)>
    'Public Structure SP_DEVINSTALL_PARAMS
    '    Public cbSize As Long
    '    Public Flags As Long
    '    Public FlagsEx As Long
    '    Public hWndParent As Long
    '    Public InstallMsgHandler As Long 'PSP_FILE_CALLBACK
    '    Public InstallMsgHandlerContext As Long 'PVOID
    '    Public FileQueue As Long 'HSPFILEQ
    '    Public ClassInstallReserved As Long 'ULONG_PTR
    '    Public Reserved As Long
    '    Public DriverPath() As Byte ' A-version
    'End Structure


    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure SP_DEVINSTALL_PARAMS
        Public cbSize As Integer
        Public Flags As Integer
        Public FlagsEx As Integer
        Public hwndParent As IntPtr
        Public InstallMsgHandler As IntPtr
        Public InstallMsgHandlerContext As IntPtr
        Public FileQueue As IntPtr
        Public ClassInstallReserved As IntPtr
        Public Reserved As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public DriverPath As String
    End Structure

    ''' <summary>
    ''' The DiInstallDevice function installs a specified driver that is preinstalled in the driver store on a specified device that is present in the system.
    ''' </summary>
    ''' <param name="hParent">A handle to the top-level window that DiInstallDevice uses to display any user interface component that is associated with installing the device. This parameter is optional and can be set to NULL.</param>
    ''' <param name="lpInfoSet">A handle to a device information set that contains a device information element that represents the specified device.</param>
    ''' <param name="DeviceInfoData">A pointer to an SP_DEVINFO_DATA structure that represents the specified device in the specified device information set.</param>
    ''' <param name="DriverInfoData">An pointer to an SP_DRVINFO_DATA structure that specifies the driver to install on the specified device. This parameter is optional and can be set to NULL. If this parameter is NULL, DiInstallDevice searches the drivers preinstalled in the driver store for the driver that is the best match to the specified device, and, if one is found, installs the driver on the specified device.</param>
    ''' <param name="Flags">A value of type DWORD that specifies zero or the following flag: DIIDFLAG_SHOWSEARCHUI DIIDFLAG_NOFINISHINSTALLUI DIIDFLAG_INSTALLNULLDRIVER DIIDFLAG_INSTALLCOPYINFDRIVERS</param>
    ''' <param name="NeedReboot">A pointer to a value of type BOOL that DiInstallDevice sets to indicate whether a system restart is required to complete the installation. This parameter is optional and can be set to NULL. If this parameter is supplied and a system restart is required to complete the installation, DiInstallDevice sets the value to TRUE. In this case, the caller is responsible for restarting the system. If this parameter is supplied and a system restart is not required, DiInstallDevice sets this parameter to FALSE. If this parameter is NULL and a system restart is required to complete the installation, DiInstallDevice displays a system restart dialog box.</param>
    ''' <returns></returns>
    <DllImport("newdev.dll", SetLastError:=True)>
    Public Shared Function DiInstallDevice(ByVal hParent As IntPtr, ByVal lpInfoSet As IntPtr, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByRef DriverInfoData As SP_DRVINFO_DATA, ByVal Flags As UInt32, ByRef NeedReboot As Boolean) As Boolean
    End Function

    ''' <summary>
    ''' The InstallSelectedDriver function is deprecated. For Windows Vista and later, use DiInstallDevice instead. You should only call InstallSelectedDriver if it is necessary to install a specific driver on a specific device.
    ''' </summary>
    ''' <param name="hWndParent"></param>
    ''' <param name="DeviceInfoSet"></param>
    ''' <param name="Reserved"></param>
    ''' <param name="Backup"></param>
    ''' <param name="bReboot"></param>
    ''' <returns></returns>
    <DllImport("newdev.dll", SetLastError:=True)>
    Public Shared Function InstallSelectedDriver(ByVal hWndParent As IntPtr, ByVal DeviceInfoSet As IntPtr, ByVal Reserved As Long, ByVal Backup As Long, ByRef bReboot As Long) As Boolean
    End Function

    ''' <summary>
    ''' The SetupDiGetClassDevs function returns a handle to a device information set that contains requested device information elements for a local computer. 
    ''' </summary>
    ''' <param name="ClassGuid">A pointer to the GUID for a device setup class or a device interface class. This pointer is optional and can be NULL. For more information about how to set ClassGuid, see the following Comments section. </param>
    ''' <param name="Enumerator">A pointer to a NULL-terminated string </param>
    ''' <param name="hwndParent">A handle to the top-level window to be used for a user interface that is associated with installing a device instance in the device information set. This handle is optional and can be NULL. </param>
    ''' <param name="Flags">A variable of type DWORD that specifies control options that filter the device information elements that are added to the device information set.</param>
    ''' <returns>If the operation succeeds, SetupDiGetClassDevs returns a handle to a device information set that contains all installed devices that matched the supplied parameters. If the operation fails, the function returns INVALID_HANDLE_VALUE. To get extended error information, call GetLastError. </returns>
    ''' <remarks>The caller of SetupDiGetClassDevs must delete the returned device information set when it is no longer needed by calling SetupDiDestroyDeviceInfoList. </remarks>
    <DllImport("setupapi.dll", EntryPoint:="SetupDiGetClassDevsW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=True, CallingConvention:=CallingConvention.Winapi)>
    Private Shared Function SetupDiGetClassDevs(
       ByRef ClassGuid As Guid,
       ByVal Enumerator As Integer,
       ByVal hwndParent As Integer,
       ByVal Flags As Integer) As IntPtr
    End Function




    ''' <summary>
    ''' The SetupDiEnumDeviceInfo function returns a SP_DEVINFO_DATA structure that specifies a device information element in a device information set. 
    ''' </summary>
    ''' <param name="DeviceInfoSet">A handle to the device information set for which to return an SP_DEVINFO_DATA structure that represents a device information element. </param>
    ''' <param name="MemberIndex">A zero-based index of the device information element to retrieve. </param>
    ''' <param name="DeviceInfoData">A pointer to an SP_DEVINFO_DATA structure to receive information about an enumerated device information element. The caller must set DeviceInfoData.cbSize to sizeof(SP_DEVINFO_DATA).</param>
    ''' <returns>The function returns TRUE if it is successful. Otherwise, it returns FALSE and the logged error can be retrieved with a call to GetLastError.</returns>
    ''' <remarks>Repeated calls to this function return a device information element for a different device. This function can be called repeatedly to get information about all devices in the device information set.</remarks>
    <DllImport("setupapi.dll", EntryPoint:="SetupDiEnumDeviceInfo", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=True, CallingConvention:=CallingConvention.Winapi)>
    Private Shared Function SetupDiEnumDeviceInfo(
           ByVal DeviceInfoSet As Integer,
           ByVal MemberIndex As Integer,
           ByRef DeviceInfoData As SP_DEVINFO_DATA) As Boolean
    End Function


    ''' <summary>
    ''' The SetupDiCreateDeviceInfoList function creates an empty device information set and optionally associates the set with a device setup class and a top-level window.
    ''' </summary>
    ''' <param name="ClassGuid">A pointer to the GUID of the device setup class to associate with the newly created device information set. If this parameter is specified, only devices of this class can be included in this device information set. If this parameter is set to NULL, the device information set is not associated with a specific device setup class. </param>
    ''' <param name="hwndParent">A handle to the top-level window to use for any user interface that is related to non-device-specific actions (such as a select-device dialog box that uses the global class driver list). This handle is optional and can be NULL. If a specific top-level window is not required, set hwndParent to NULL. </param>
    ''' <returns>The function returns a handle to an empty device information set if it is successful. Otherwise, it returns INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.</returns>
    ''' <remarks>The caller of this function must delete the returned device information set when it is no longer needed by calling SetupDiDestroyDeviceInfoList. </remarks>
    <DllImport("Setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiCreateDeviceInfoList(ByRef ClassGuid As Guid, ByVal hwndParent As IntPtr) As IntPtr
    End Function



    '  HKEY
    'SetupDiOpenDevRegKey(
    '  IN HDEVINFO  DeviceInfoSet,
    '  IN PSP_DEVINFO_DATA  DeviceInfoData,
    '  IN DWORD  Scope,
    '  IN DWORD  HwProfile,
    '  IN DWORD  KeyType,
    '  IN REGSAM  samDesired
    '  );
    ''' <summary>
    ''' The SetupDiOpenDevRegKey function opens a registry key for device-specific configuration information.
    ''' </summary>
    ''' <param name="hDeviceInfoSet">A handle to the device information set that contains a device information element that represents the device for which to open a registry key.</param>
    ''' <param name="deviceInfoData">A pointer to an SP_DEVINFO_DATA structure that specifies the device information element in DeviceInfoSet. </param>
    ''' <param name="scope">The scope of the registry key to open. The scope determines where the information is stored. The scope can be global or specific to a hardware profile. The scope is specified by one of the following values: DICS_FLAG_GLOBAL or DICS_FLAG_CONFIGSPECIFIC </param>
    ''' <param name="hwProfile">A hardware profile value, which is set as follows:If Scope is set to DICS_FLAG_CONFIGSPECIFIC, HwProfile specifies the hardware profile of the key that is to be opened. If HwProfile is 0, the key for the current hardware profile is opened. If Scope is DICS_FLAG_GLOBAL, HwProfile is ignored.  </param>
    ''' <param name="parameterRegistryValueKind">The type of registry storage key to open, which can be one of the following values: DIREG_DEV Open a hardware key for the device. DIREG_DRV Open a software key for the device. </param>
    ''' <param name="samDesired">The registry security access that is required for the requested key. For information about registry security access values of type REGSAM, see the Microsoft Windows SDK documentation. </param>
    ''' <returns>If the function is successful, it returns a handle to an opened registry key where private configuration data about this device instance can be stored/retrieved. If the function fails, it returns INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.</returns>
    ''' <remarks>Depending on the value that is passed in the samDesired parameter, it might be necessary for the caller of this function to be a member of the Administrators group.Close the handle returned from this function by calling RegCloseKey.The specified device instance must be registered before this function is called. However, be aware that the operating system automatically registers PnP device instances. For information about how to register non-PnP device instances, see SetupDiRegisterDeviceInfo.</remarks>
    <DllImport("Setupapi", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiOpenDevRegKey(ByVal hDeviceInfoSet As IntPtr, ByRef deviceInfoData As SP_DEVINFO_DATA, ByVal scope As Integer, ByVal hwProfile As Integer,
                                                ByVal parameterRegistryValueKind As Integer, ByVal samDesired As Integer) As IntPtr
    End Function

    ''' <summary>
    ''' The SetupDiCallClassInstaller function calls the appropriate class installer, and any registered co-installers, with the specified installation request (DIF code).
    ''' </summary>
    ''' <param name="InstallFunction">The device installation request (DIF request) to pass to the co-installers and class installer. DIF codes have the format DIF_XXX and are defined in Setupapi.h. </param>
    ''' <param name="DeviceInfoSet">A handle to a device information set for the local computer. This set contains a device installation element which represents the device for which to perform the specified installation function. </param>
    ''' <param name="DeviceInfoData">A pointer to an SP_DEVINFO_DATA structure that specifies the device information element in the DeviceInfoSet that represents the device for which to perform the specified installation function. This parameter is optional and can be set to NULL.</param>
    ''' <returns>The function returns TRUE if it is successful. Otherwise, it returns FALSE and the logged error can be retrieved by making a call to GetLastError.</returns>
    ''' <remarks></remarks>
    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiCallClassInstaller(ByVal InstallFunction As UInt32, ByVal DeviceInfoSet As IntPtr, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Boolean
    End Function


    ''' <summary>
    ''' The SetupDiRegisterDeviceInfo function is the default handler for the DIF_REGISTERDEVICE request. 
    ''' </summary>
    ''' <param name="DeviceInfoSet">A handle to a device information set that contains a device information element that represents the device to register. The device information set must not contain any remote elements. </param>
    ''' <param name="deviceInfoData">A pointer to an SP_DEVINFO_DATA structure that specifies the device information element in DeviceInfoSet. This is an IN-OUT parameter because DeviceInfoData.DevInst might be updated with a new handle value upon return. </param>
    ''' <param name="Flags">A flag value that controls how the device is registered, which can be zero or the following value: SPRDI_FIND_DUPS. Search for a previously-existing device instance that corresponds to the device that is represented by DeviceInfoData. If this flag is not specified, the device instance is registered regardless of whether a device instance already exists for it.If the caller supplies CompareProc, the caller must also set this flag.</param>
    ''' <param name="CompareProc">A pointer to a comparison callback function to use in duplicate detection. This parameter is optional and can be NULL. </param>
    ''' <param name="CompareContext">A pointer to a caller-supplied context buffer that is passed into the callback function. This parameter is ignored if CompareProc is not specified. </param>
    ''' <param name="DupDeviceInfoData">A pointer to an SP_DEVINFO_DATA structure to receive information about a duplicate device instance, </param>
    ''' <returns>The function returns TRUE if it is successful. Otherwise, it returns FALSE and the logged error can be retrieved with a call to GetLastError.</returns>
    ''' <remarks></remarks>
    <DllImport("Setupapi.dll", EntryPoint:="SetupDiRegisterDeviceInfo", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiRegisterDeviceInfo(ByVal DeviceInfoSet As IntPtr, ByRef deviceInfoData As SP_DEVINFO_DATA, ByVal Flags As Integer,
                            ByVal CompareProc As IntPtr, ByVal CompareContext As IntPtr, ByVal DupDeviceInfoData As IntPtr) As Boolean
    End Function


    Public Const DI_NEEDREBOOT As Integer = &H100
    Public Const DI_NEEDRESTART As Integer = &H80

    ''' <summary>
    ''' The SetupDiGetSelectedDriver function retrieves the selected driver for a device information set or a particular device information element.
    ''' </summary>
    ''' <param name="DeviceInfoSet">A handle to the device information set for which to retrieve a selected driver. </param>
    ''' <param name="deviceInfoData">A pointer to an SP_DEVINFO_DATA structure that specifies a device information element that represents the device in DeviceInfoSet for which to retrieve the selected driver. This parameter is optional and can be NULL. If this parameter is specified, SetupDiGetSelectedDriver retrieves the selected driver for the specified device. If this parameter is NULL, SetupDiGetSelectedDriver retrieves the selected class driver in the global class driver list that is associated with DeviceInfoSet. </param>
    ''' <param name="DriverInfoData">A pointer to an SP_DRVINFO_DATA structure that receives information about the selected driver.</param>
    ''' <returns>The function returns TRUE if it is successful. Otherwise, it returns FALSE and the logged error can be retrieved with a call to GetLastError. If a driver has not been selected for the specified device instance, the logged error is ERROR_NO_DRIVER_SELECTED.</returns>
    ''' <remarks></remarks>
    <DllImport("Setupapi.dll", EntryPoint:="SetupDiGetSelectedDriver", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiGetSelectedDriver(ByVal DeviceInfoSet As IntPtr, ByRef deviceInfoData As SP_DEVINFO_DATA, ByRef DriverInfoData As SP_DRVINFO_DATA) As Boolean
    End Function

    'HKEY SetupDiCreateDevRegKey( _In_     HDEVINFO         DeviceInfoSet,  _In_     PSP_DEVINFO_DATA DeviceInfoData,  _In_     DWORD            Scope,  _In_     DWORD            HwProfile,
    '_In_     DWORD            KeyType,  _In_opt_ HINF             InfHandle,  _In_opt_ PCTSTR           InfSectionName);

    <DllImport("Setupapi.dll", EntryPoint:="SetupDiCreateDevRegKey", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiCreateDevRegKey(ByVal DeviceInfoSet As IntPtr, ByRef deviceInfoData As SP_DEVINFO_DATA, ByVal Scope As Integer, ByVal HwProfile As Integer,
                               ByVal KeyType As Integer, ByVal InfHandle As IntPtr, ByVal InfSectionName As IntPtr) As IntPtr
    End Function


    '    BOOL SetupDiGetINFClass( _In_   PCTSTR InfName,  _Out_     LPGUID ClassGuid,  _Out_     PTSTR  ClassName,  _In_  DWORD  ClassNameSize,  _Out_opt_ PDWORD RequiredSize);
    <DllImport("setupapi.dll", EntryPoint:="SetupDiGetINFClass", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiGetINFClass(ByVal InfName As String, ByRef ClassGuid As Guid, ByVal ClassName As StringBuilder, ByVal ClassNameSize As Integer, Optional ByRef RequiredSize As UInteger = 0) As Boolean
    End Function
    ' Public Declare Auto Function SetupDiGetINFClass Lib "setupapi.dll" Alias "SetupDiGetINFClassW" (ByVal InfName As String, ByRef ClassGuid As Guid, ByVal ClassName As String, ByVal ClassNameSize As Integer, ByRef RequiredSize As Integer) As Boolean


    ' BOOL WINAPI SetupEnumInfSections(  _In_  HINF InfHandle,  _In_ UINT EnumerationIndex,  _Out_opt_  PTSTR ReturnBuffer,  _In_ DWORD ReturnBufferSize,  _Out_opt_  PDWORD RequiredSize);
    <DllImport("setupapi.dll", EntryPoint:="SetupEnumInfSections", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupEnumInfSections(ByVal infHandle As IntPtr, ByVal index As UInteger, ByVal ReturnBuffer As StringBuilder, ByVal ReturnBufferSize As Integer, ByRef RequiredSize As UInteger) As Integer
    End Function


    <DllImport("Setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiCreateDeviceInfo(ByVal DeviceInfoSet As IntPtr, ByVal DeviceName As [String], ByRef ClassGuid As Guid, ByVal DeviceDescription As String, ByVal hwndParent As IntPtr, ByVal CreationFlags As Int32,
 ByRef DeviceInfoData As SP_DEVINFO_DATA) As Boolean
    End Function

    <DllImport("setupapi.dll",
   EntryPoint:="SetupDiDestroyDeviceInfoList", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True, PreserveSig:=True, CallingConvention:=CallingConvention.Winapi)>
    Private Shared Function SetupDiDestroyDeviceInfoList(ByVal DeviceInfoSet As Integer) As Boolean
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiGetClassDevsA(ByRef ClassGuid As Guid, ByVal Enumerator As UInt32, ByVal hwndParent As IntPtr, ByVal Flags As UInt32) As IntPtr
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupDiGetDeviceRegistryProperty(
    ByVal DeviceInfoSet As Integer,
    ByRef DeviceInfoData As SP_DEVINFO_DATA,
    ByVal [Property] As Integer,
    ByRef PropertyRegDataType As Integer,
    ByVal PropertyBuffer As Byte(),
    ByVal PropertyBufferSize As Integer,
    ByRef RequiredSize As Integer) As Boolean
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetupDiSetDeviceRegistryProperty(
    ByVal DeviceInfoSet As IntPtr,
    ByRef DeviceInfoData As SP_DEVINFO_DATA,
    ByVal [Property] As UInteger, ByVal PropertyBuffer As String, ByVal PropertyBufferSize As Integer) As Boolean
    End Function


    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SetupCopyOEMInf(ByVal SourceInfFileName As String, ByVal OEMSourceMediaLocation As String, ByVal OEMSourceMediaType As Integer, ByVal CopyStyle As Integer, ByVal DestinationInfFileName As String, ByVal DestinationInfFileNameSize As Integer,
    ByVal RequiredSize As Integer, ByVal DestinationInfFileNameComponent As String) As Boolean

    End Function


    'LONG WINAPI RegQueryValueEx(  _In_ HKEY    hKey, _In_opt_  LPCTSTR lpValueName, _Reserved_  LPDWORD lpReserved, _Out_opt_   LPDWORD lpType,  _Out_opt_   LPBYTE  lpData,  _Inout_opt_ LPDWORD lpcbData);
    ''' <summary>
    ''' Retrieves the type and data for the specified value name associated with an open registry key.To ensure that any string values (REG_SZ, REG_MULTI_SZ, and REG_EXPAND_SZ) returned are null-terminated, use the RegGetValue function.
    ''' </summary>
    ''' <param name="hKey">A handle to an open registry key. The key must have been opened with the KEY_QUERY_VALUE access right. This handle is returned by the RegCreateKeyEx.
    '''It can also be one of the following predefined keys: HKEY_CLASSES_ROOT, HKEY_CURRENT_CONFIG, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_PERFORMANCE_DATA, HKEY_PERFORMANCE_NLSTEXT, HKEY_PERFORMANCE_TEXT, HKEY_USERS
    ''' </param>
    ''' <param name="lpValueName">The name of the registry value. If lpValueName specifies a value that is not in the registry, the function returns ERROR_FILE_NOT_FOUND.</param>
    ''' <param name="lpReserved">This parameter is reserved and must be NULL.</param>
    ''' <param name="lpType">A pointer to a variable that receives a code indicating the type of data stored in the specified value. For a list of the possible type codes, see Registry Value Types. The lpType parameter can be NULL if the type code is not required.</param>
    ''' <param name="lpData">A pointer to a buffer that receives the value's data. This parameter can be NULL if the data is not required.</param>
    ''' <param name="lpcbData">A pointer to a variable that specifies the size of the buffer pointed to by the lpData parameter, in bytes. When the function returns, this variable contains the size of the data copied to lpData.The lpcbData parameter can be NULL only if lpData is NULL.
    ''' If the data has the REG_SZ, REG_MULTI_SZ or REG_EXPAND_SZ type, this size includes any terminating null character or characters unless the data was stored without them. For more information, see Remarks. If the buffer specified by lpData parameter is not large enough to hold the data, the function returns ERROR_MORE_DATA and stores the required buffer size in the variable pointed to by lpcbData. In this case, the contents of the lpData buffer are undefined.
    ''' If lpData is NULL, and lpcbData is non-NULL, the function returns ERROR_SUCCESS and stores the size of the data, in bytes, in the variable pointed to by lpcbData. This enables an application to determine the best way to allocate a buffer for the value's data.</param>
    ''' <returns>If the function succeeds, the return value is ERROR_SUCCESS. If the function fails, the return value is a system error code. If the lpData buffer is too small to receive the data, the function returns ERROR_MORE_DATA. f the lpValueName registry value does not exist, the function returns ERROR_FILE_NOT_FOUND. </returns>
    ''' <remarks></remarks>
    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Friend Shared Function RegQueryValueEx(ByVal hKey As IntPtr, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As System.Text.StringBuilder, ByRef lpcbData As Integer) As Integer
    End Function


    '    LONG WINAPI RegSetValueEx(
    '  _In_             HKEY    hKey,
    '  _In_opt_         LPCTSTR lpValueName,
    '  _Reserved_       DWORD   Reserved,
    '  _In_             DWORD   dwType,
    '  _In_       const BYTE    *lpData,
    '  _In_             DWORD   cbData
    ');
    ''' <summary>
    ''' Sets the data and type of a specified value under a registry key.
    ''' </summary>
    ''' <param name="hKey">A handle to an open registry key. The key must have been opened with the KEY_SET_VALUE access right. </param>
    ''' <param name="lpValueName">The name of the value to be set. If a value with this name is not already present in the key, the function adds it to the key.If lpValueName is NULL or an empty string, "", the function sets the type and data for the key's unnamed or default value.</param>
    ''' <param name="lpReserved">This parameter is reserved and must be zero. </param>
    ''' <param name="dwType">The type of data pointed to by the lpData parameter. For a list of the possible types, see Registry Value Types.</param>
    ''' <param name="lpData">The data to be stored. For string-based types, such as REG_SZ, the string must be null-terminated. With the REG_MULTI_SZ data type, the string must be terminated with two null characters. String-literal values must be formatted using a backslash preceded by another backslash as an escape character. For example, specify "C:\\mydir\\myfile" to store the string "C:\mydir\myfile".</param>
    ''' <param name="cbData">The size of the information pointed to by the lpData parameter, in bytes. If the data is of type REG_SZ, REG_EXPAND_SZ, or REG_MULTI_SZ, cbData must include the size of the terminating null character or characters.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Friend Shared Function RegSetValueEx(
    ByVal hKey As IntPtr,
    ByVal lpValueName As String,
    ByVal lpReserved As Integer,
    ByVal dwType As Integer,
    ByVal lpData As String,
    ByVal cbData As Integer) As Integer
    End Function

    ''' <summary>
    ''' Closes a handle to the specified registry key.
    ''' </summary>
    ''' <param name="hKey">A handle to the open key to be closed. The handle must have been opened by the RegCreateKeyEx, RegCreateKeyTransacted, RegOpenKeyEx, RegOpenKeyTransacted, or RegConnectRegistry function.</param>
    ''' <returns>If the function succeeds, the return value is ERROR_SUCCESS.If the function fails, the return value is a nonzero error code defined in Winerror.h. You can use the FormatMessage function with the FORMAT_MESSAGE_FROM_SYSTEM flag to get a generic description of the error.</returns>
    ''' <remarks>The handle for a specified key should not be used after it has been closed, because it will no longer be valid. Key handles should not be left open any longer than necessary.The RegCloseKey function does not necessarily write information to the registry before returning; it can take as much as several seconds for the cache to be flushed to the hard disk. If an application must explicitly write registry information to the hard disk, it can use the RegFlushKey function. RegFlushKey, however, uses many system resources and should be called only when necessary.</remarks>
    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function RegCloseKey(ByVal hKey As IntPtr) As Integer
    End Function


    Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Integer = &H100
    Const FORMAT_MESSAGE_FROM_SYSTEM As Integer = &H1000
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dwFlags"></param>
    ''' <param name="lpSource"></param>
    ''' <param name="dwMessageId"></param>
    ''' <param name="dwLanguageId"></param>
    ''' <param name="lpBuffer"></param>
    ''' <param name="nSize"></param>
    ''' <param name="Arguments"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("Kernel32.dll", EntryPoint:="FormatMessageW", SetLastError:=True, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByRef lpBuffer As [String], ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    End Function

    Public Shared Function FormatErrorMessage(ByVal errore As Integer) As String
        Try
            Dim h As Integer = -1
            Dim dwLanguageId As Integer = 0
            Dim lpBuffer As String = ""

            dwLanguageId = System.Globalization.CultureInfo.CurrentCulture.LCID
            ' dwLanguageId = System.Globalization.CultureInfo.GetCultureInfo("en-GB").LCID
            h = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM, IntPtr.Zero, errore, dwLanguageId, lpBuffer, 0, IntPtr.Zero)

            '  FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,    NULL,   Err,     MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),   (LPTSTR) &lpMessageBuffer,  
            ' 0,       NULL ))

            If h <> -1 AndAlso lpBuffer <> "" Then
                Return String.Format("Error: 0x{0:x8}", errore) + ", " + lpBuffer
            Else
                Return String.Format("Error: 0x{0:x8}", errore) + ", " + "unknown"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    '<DllImport("newdev.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    'Public Shared Function UpdateDriverForPlugAndPlayDevices(<[In](), [Optional]()> ByVal hwndParent As IntPtr, <[In]()> ByVal HardwareId As String, <[In]()> ByVal FullInfPath As String, <[In]()> ByVal InstallFlags As UInteger, <Out(), [Optional]()> ByVal bRebootRequired As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
    'End Function

    ''' <summary>
    ''' Given an INF file and a hardware ID, the UpdateDriverForPlugAndPlayDevices function installs updated drivers for devices that match the hardware ID.
    ''' </summary>
    ''' <param name="hwndParent"></param>
    ''' <param name="HardwareId"></param>
    ''' <param name="FullInfPath"></param>
    ''' <param name="InstallFlags"></param>
    ''' <param name="bRebootRequired"></param>
    ''' <returns></returns>
    <DllImport("newdev.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function UpdateDriverForPlugAndPlayDevices(
             ByVal hwndParent As IntPtr,
             ByVal HardwareId As String,
             ByVal FullInfPath As String,
             ByVal InstallFlags As UInteger,
             ByVal bRebootRequired As Boolean) As Boolean
    End Function


    ''' <summary>
    ''' The DiInstallDriver function preinstalls a driver in the driver store and then installs the driver on devices present in the system that the driver supports.
    ''' </summary>
    ''' <param name="hwndParent"></param>
    ''' <param name="FullInfPath"></param>
    ''' <param name="InstallFlags"></param>
    ''' <param name="bRebootRequired"></param>
    ''' <returns></returns>
    <DllImport("newdev.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function DiInstallDriver(
             ByVal hwndParent As IntPtr,
             ByVal FullInfPath As String,
             ByVal InstallFlags As UInteger,
             ByVal bRebootRequired As Boolean) As Boolean
    End Function


    Private CLASS_GUID_MODEM As Guid = New Guid("{4D36E96D-E325-11CE-BFC1-08002BE10318}") ' modem
    '{4d36e96d-e325-11ce-bfc1-08002be10318}
    Private CLASS_GUID_PORTS As Guid = New Guid("{4D36E978-E325-11CE-BFC1-08002BE10318}") ' ports

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure INSTALLERINFO
        Public pApplicationId As String
        Public pDisplayName As String
        Public pProductName As String
        Public pMfgName As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SP_DEVINFO_DATA
        Public cbSize As Integer
        Public classGuid As Guid
        Public propertyId As Integer
        Public reserved As IntPtr
    End Structure



    ''' <summary>
    ''' An SP_DRVINFO_DATA structure contains information about a driver.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SP_DRVINFO_DATA
        Public cbSize As Integer
        Public DriverType As Integer
        Private Reserved As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public Description As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public MfgName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public ProviderName As String
    End Structure

#End Region

#Region "Costanti"



    'private Const HardwareId As String = "PNPC031"
    Private Const HardwareId_HS As String = "M5PNPHS"
    Private Const HardwareId_LS As String = "M5PNPLS"

    Const DIGCF_DEFAULT As Int32 = &H1 'Return only the device that is associated with the system default device interface, if one is set, for the specified device interface classes. 
    Const DIGCF_PRESENT As Int32 = &H2 'Return only devices that are currently present in a system. 
    Const DIGCF_ALLCLASSES As Int32 = &H4 'Return a list of installed devices for all device setup classes or all device interface classes. 
    Const DIGCF_PROFILE As Int32 = &H8 'Return only devices that are a part of the current hardware profile. 
    Const DIGCF_DEVICEINTERFACE As Int32 = &H10 'Return devices that support device interfaces for the specified device interface classes. This flag must be set in the Flags parameter if the Enumerator parameter specifies a device instance ID. 


    Const DRIVER_PACKAGE_REPAIR As Int32 = &H1
    Const DRIVER_PACKAGE_SILENT As Int32 = &H2
    Const DRIVER_PACKAGE_FORCE As Int32 = &H4
    Const DRIVER_PACKAGE_ONLY_IF_DEVICE_PRESENT As Int32 = &H8
    Const DRIVER_PACKAGE_LEGACY_MODE As Int32 = &H10
    Const DRIVER_PACKAGE_DELETE_FILES As Int32 = &H20

    Const INVALID_HANDLE_VALUE As Int64 = -1




    Const KEY_ALL_ACCESS As Integer = &HF003F
    ' Permission For all types Of access.
    Const KEY_CREATE_LINK As Integer = &H20
    ' Permission To create symbolic links.
    Const KEY_CREATE_SUB_KEY As Integer = &H4
    'Permission to create subkeys.
    Const KEY_ENUMERATE_SUB_KEYS As Integer = &H8
    'Permission to enumerate subkeys.
    Const KEY_EXECUTE As Integer = &H20019
    'Same as KEY_READ.
    Const KEY_NOTIFY As Integer = &H10
    'Permission to give change notification.
    Const KEY_QUERY_VALUE As Integer = &H1
    'Permission to query subkey data.
    Const KEY_READ As Integer = &H20019
    'Permission for general read access.
    Const KEY_SET_VALUE As Integer = &H2
    'Permission to set subkey data.
    Const KEY_WRITE As Integer = &H20006
    'Permission for general write access

    Const MAX__DESC As Integer = 256

    Const DICS_FLAG_GLOBAL As Int32 = &H1 'Open a key to store global configuration information. This information is not specific to a particular hardware profile. This opens a key that is rooted at HKEY_LOCAL_MACHINE. The exact key opened depends on the value of the KeyType parameter. 

    Const DIREG_DEV As Integer = &H1  '#define DIREG_DEV       0x00000001          // Open/Create/Delete device hardware key  
    Const DIREG_DRV As Integer = &H2  '#define DIREG_DRV       0x00000002          // Open/Create/Delete driver software key
    Const DIREG_BOTH As Integer = &H4 '#define DIREG_BOTH      0x00000004          // Delete both driver and Device key

    Const ERROR_SUCCESS As Integer = 0
    Const ERROR_INVALID_DATA As Integer = 13
    Const ERROR_INSUFFICIENT_BUFFER As Integer = 122
    Const ERROR_NO_MORE_ITEMS As Integer = 259
    Const ERROR_IN_WOW64 As Integer = &H235 '#define ERROR_IN_WOW64 (APPLICATION_ERROR_MASK|ERROR_SEVERITY_ERROR|0x235)
    Const ERROR_KEY_DOES_NOT_EXIST As Integer = -536870396 '#define ERROR_KEY_DOES_NOT_EXIST  (APPLICATION_ERROR_MASK|ERROR_SEVERITY_ERROR|0x204)

    Const SPDRP_HARDWAREID As Integer = &H1 ' HardwareID (R/W)
    Const SPDRP_FRIENDLYNAME As Integer = 12

    Const INSTALLFLAG_FORCE As Integer = &H1
    Const INSTALLFLAG_READONLY As Integer = &H2
    Const INSTALLFLAG_NONINTERACTIVE As Integer = &H4

    Const DICD_GENERATE_ID As Integer = &H1
    Const DIF_REMOVE As Integer = &H5 '#define DIF_REMOVE 0x00000005 'A DIF_REMOVE request notifies an installer that Setup is about to remove a device and gives the installer an opportunity to prepare for the removal when Sent


#End Region
    Enum EnumM5Speed
        HighSpeed
        LowSpeed
    End Enum

    Public Shared Function GetHardawrareID(speed As EnumM5Speed) As String
        Try
            Select Case speed
                Case EnumM5Speed.LowSpeed
                    Return HardwareId_LS
                Case EnumM5Speed.HighSpeed
                    Return HardwareId_HS
                Case Else
                    Return HardwareId_LS
            End Select
        Catch ex As Exception

        End Try
    End Function



    Public Sub PrintLog(s As String)
        Try
            _writeLog(s)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ResetLog()
        Try
            Safe.Control.Text("", Me.txtLog)

            If PathLog IsNot Nothing AndAlso PathLog.Trim <> "" Then

                If Not My.Computer.FileSystem.DirectoryExists(PathLog) Then
                    My.Computer.FileSystem.CreateDirectory(PathLog)
                End If

                Dim path As String = My.Computer.FileSystem.CombinePath(PathLog, nomeFileLog)

                My.Computer.FileSystem.WriteAllText(path, "", False)

            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall" + vbCrLf + "Path: " + PathLog + vbCrLf + "File: " + nomeFileLog, "ResetLog", ex, True)
        End Try
    End Sub

    Public Sub WriteLine(ByVal desc As String, ByVal EventDescription As String, ByVal ErrorCode As Int32)
        Try
            Dim s As String = String.Format(desc, EventDescription, ErrorCode)

            Safe.Textbox.AppendText(s + vbCrLf, Me.txtLog)

            _writeLog(s)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "WriteLine", ex, True)
        End Try

    End Sub

    Public Sub WriteLine(ByVal desc As String)
        Try
            Safe.Textbox.AppendText(desc + vbCrLf, Me.txtLog)

            _writeLog(desc)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "WriteLine", ex, True)
        End Try
    End Sub

    Private Sub _writeLog(s As String)
        Try
            If PathLog IsNot Nothing AndAlso PathLog.Trim <> "" Then
                My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.CombinePath(PathLog, nomeFileLog), "[" + Now.Date.ToShortDateString + " " + Format(Now, "hh:mm:ss.fff") + "] " + s.Replace(vbCrLf, ". ") + vbCrLf, True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Check if port is plugged
    ''' </summary>
    ''' <param name="PortName"></param>
    ''' <param name="DisableAndEnableUsb"></param>
    ''' <returns></returns>
    Public Function PortExist(ByVal PortName As String, ByVal DisableAndEnableUsb As Boolean) As Boolean
        Try
            Dim match As Boolean = False

            WriteLine("PortExist?")

            WriteLine("SetupDiGetClassDevs ..." + CLASS_GUID_PORTS.ToString.ToUpper)


            ' The SetupDiGetClassDevs function returns a handle to a device information set that contains requested device information elements for a local computer. 
            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_PORTS, 0, 0, DIGCF_PRESENT)

            If hDeviceInfoSet <> New IntPtr(-1) AndAlso hDeviceInfoSet <> IntPtr.Zero Then

                Dim devInfoElem As SP_DEVINFO_DATA
                Dim index As Integer = 0
                devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

                While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True
                    index += 1
                    Dim hKeyDev As System.IntPtr

                    ' Opens a registry key for device-specific configuration information.
                    hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ)
                    Dim szDevDesc As New System.Text.StringBuilder(20, MAX__DESC)

                    If hKeyDev <> New IntPtr(-1) Then

                        Dim esito As Integer = RegQueryValueEx(hKeyDev, "Portname", Nothing, Nothing, szDevDesc, MAX__DESC)
                        If (ERROR_SUCCESS = esito) Then
                            RegCloseKey(hKeyDev)
                            Dim DevDesc As String = szDevDesc.ToString

                            WriteLine(String.Format("PortExist.Comparing :  ""{0}"" with ""{1}""", PortName, DevDesc))

                            If (PortName = DevDesc) Then

                                WriteLine(String.Format("PortExist.Found :  ""{0}"" = ""{1}""", PortName, DevDesc))

                                'Dim p As New HardwareHelperLib.HH_Lib
                                'p.ResetDevice(hDeviceInfoSet, Nothing)
                                If DisableAndEnableUsb Then


                                    ClassDevice.Disable(PortName)

                                    wait(1000)

                                    ClassDevice.Enable(PortName)

                                    wait(1000)

                                End If

                                match = True
                                Exit While
                            End If

                        End If
                    End If

                End While

                SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            End If
            If match = False Then WriteLine(String.Format("PortExist.Not Found :  ""{0}""", PortName))
            Return match

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "PortExist", ex, True)
        End Try

    End Function

    Public Function ModemExist(ByVal PortName As String) As Boolean
        Try
            Dim match As Boolean = False

            WriteLine("Modem exist?")

            WriteLine("SetupDiGetClassDevs ..." + CLASS_GUID_PORTS.ToString.ToUpper)

            ' The SetupDiGetClassDevs function returns a handle to a device information set that contains requested device information elements for a local computer. 
            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_MODEM, 0, 0, DIGCF_PRESENT)

            If hDeviceInfoSet <> New IntPtr(-1) AndAlso hDeviceInfoSet <> IntPtr.Zero Then

                Dim devInfoElem As SP_DEVINFO_DATA
                Dim index As Integer = 0
                devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

                While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True
                    index += 1
                    Dim hKeyDev As System.IntPtr

                    ' Opens a registry key for device-specific configuration information.
                    hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DRV, KEY_READ) 'DIREG_DEV
                    Dim szDevDesc As New System.Text.StringBuilder(20, MAX__DESC)

                    If hKeyDev <> New IntPtr(-1) Then

                        Dim esito As Integer = RegQueryValueEx(hKeyDev, "AttachedTo", Nothing, Nothing, szDevDesc, MAX__DESC)
                        If (ERROR_SUCCESS = esito) Then
                            RegCloseKey(hKeyDev)
                            Dim DevDesc As String = szDevDesc.ToString

                            WriteLine(String.Format("Comparing :  ""{0}"" with ""{1}""", PortName, DevDesc))

                            If (PortName = DevDesc) Then

                                WriteLine(String.Format("Found :  ""{0}"" = ""{1}""", PortName, DevDesc))

                                match = True
                                Exit While
                            End If

                        End If
                    Else
                        MessageBox.Show(New Win32Exception(Marshal.GetLastWin32Error).Message, "SetupDiOpenDevRegKey", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                End While

                SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            End If
            If match = False Then WriteLine(String.Format("Not Found :  ""{0}""", PortName))
            Return match

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "PortExist", ex, True)
        End Try

    End Function


    Public Function UninstallModem(ByVal PortName As String) As Boolean
        Try
            Dim match As Boolean = False

            WriteLine("SetupDiGetClassDevs ..." + CLASS_GUID_PORTS.ToString.ToUpper)

            ' The SetupDiGetClassDevs function returns a handle to a device information set that contains requested device information elements for a local computer. 
            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_MODEM, 0, 0, DIGCF_PRESENT)

            If hDeviceInfoSet <> New IntPtr(-1) AndAlso hDeviceInfoSet <> IntPtr.Zero Then

                Dim devInfoElem As SP_DEVINFO_DATA
                Dim index As Integer = 0
                devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

                While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True
                    index += 1
                    Dim hKeyDev As System.IntPtr

                    ' Opens a registry key for device-specific configuration information.
                    hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DRV, KEY_READ) 'DIREG_DEV
                    Dim szDevDesc As New System.Text.StringBuilder(20, MAX__DESC)

                    If hKeyDev <> New IntPtr(-1) Then

                        Dim esito As Integer = RegQueryValueEx(hKeyDev, "AttachedTo", Nothing, Nothing, szDevDesc, MAX__DESC)
                        If (ERROR_SUCCESS = esito) Then
                            RegCloseKey(hKeyDev)
                            Dim DevDesc As String = szDevDesc.ToString

                            WriteLine(String.Format("Comparing :  ""{0}"" with ""{1}""", PortName, DevDesc))

                            If (PortName = DevDesc) Then

                                WriteLine(String.Format("Found :  ""{0}"" = ""{1}""", PortName, DevDesc))

                                match = True

                                ' invoke class installer method with DIF_REMOVE flag
                                If Not SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, devInfoElem) Then
                                    DisplayError("SetupDiCallClassInstaller", "")
                                    ' failed to uninstall device
                                    SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

                                End If

                                Exit While

                            End If

                        End If
                    End If

                End While

                SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            End If
            If match = False Then WriteLine(String.Format("Not Found :  ""{0}""", PortName))
            Return match

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "PortExist", ex, True)
        End Try

    End Function

    Public Function UninstallPort(ByVal PortName As String) As Boolean
        Try
            Dim match As Boolean = False

            WriteLine("SetupDiGetClassDevs ..." + CLASS_GUID_PORTS.ToString.ToUpper)

            ' The SetupDiGetClassDevs function returns a handle to a device information set that contains requested device information elements for a local computer. 
            Dim hDeviceInfoSet As IntPtr = SetupDiGetClassDevs(CLASS_GUID_PORTS, 0, 0, DIGCF_PROFILE)

            If hDeviceInfoSet <> New IntPtr(-1) AndAlso hDeviceInfoSet <> IntPtr.Zero Then

                Dim devInfoElem As SP_DEVINFO_DATA
                Dim index As Integer = 0
                devInfoElem.cbSize = Marshal.SizeOf(devInfoElem)

                While SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, index, devInfoElem) = True
                    index += 1
                    Dim hKeyDev As System.IntPtr

                    ' Opens a registry key for device-specific configuration information.
                    hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ) 'DIREG_DRV
                    Dim szDevDesc As New System.Text.StringBuilder(20, MAX__DESC)

                    If hKeyDev <> New IntPtr(-1) Then

                        Dim esito As Integer = RegQueryValueEx(hKeyDev, "Portname", Nothing, Nothing, szDevDesc, MAX__DESC)
                        If (ERROR_SUCCESS = esito) Then
                            RegCloseKey(hKeyDev)
                            Dim DevDesc As String = szDevDesc.ToString

                            WriteLine(String.Format("Comparing :  ""{0}"" with ""{1}""", PortName, DevDesc))

                            If (PortName = DevDesc) Then

                                WriteLine(String.Format("Found :  ""{0}"" = ""{1}""", PortName, DevDesc))

                                match = True

                                ' invoke class installer method with DIF_REMOVE flag
                                If Not SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, devInfoElem) Then
                                    DisplayError("SetupDiCallClassInstaller", "")
                                    ' failed to uninstall device
                                    SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

                                End If

                                Exit While

                            End If

                        End If
                    End If

                End While

                SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            End If
            If match = False Then WriteLine(String.Format("Not Found :  ""{0}""", PortName))
            Return match

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "PortExist", ex, True)
        End Try

    End Function

    Public Sub wait(ByVal ms As Integer)
        Try
            Dim sw As New Stopwatch
            sw.Reset()
            sw.Start()
            Do

                If sw.Elapsed.TotalMilliseconds >= ms Then
                    Exit Do
                End If

                System.Threading.Thread.Sleep(100)

            Loop
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "wait", ex, True)
        End Try
    End Sub

    Public Function FindExistingDevice(ByVal HardwareId As String, ByRef lastErr As Integer) As Boolean

        Dim hDeviceInfoSet As System.IntPtr

        Try
            lastErr = 0
            'Setupapi.h (include Setupapi.h). Available in Microsoft Windows 2000 and later versions of Windows.

            ' Create a Device Information Set with all present devices.
            hDeviceInfoSet = SetupDiGetClassDevs(Nothing, Nothing, Nothing, DIGCF_ALLCLASSES Or DIGCF_PRESENT)

            If (hDeviceInfoSet = New IntPtr(-1)) Then
                MessageBox.Show("GetClassDevs(All Present Devices)" + vbCrLf + Marshal.GetLastWin32Error.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            WriteLine(String.Format("Search for Device ID: {0} ...", HardwareId))

            ' Enumerate through all Devices.
            Dim Found As Boolean = False
            Dim i As Integer = 0
            Dim errore As Integer = 0
            Dim DataType As Integer = 0
            Dim DeviceInfoData As New SP_DEVINFO_DATA
            Dim buffer(-1) As Byte
            Dim buffersize As Integer = 0

            'DeviceInfoData.cbSize = sizeof(SP_DEVINFO_DATA)
            Dim esitoEnumDeviceInfoOk As Boolean = False
            DeviceInfoData.cbSize = Marshal.SizeOf(DeviceInfoData)

            While Not Found And SetupDiEnumDeviceInfo(hDeviceInfoSet.ToInt32, i, DeviceInfoData)
                i += 1

                While Not SetupDiGetDeviceRegistryProperty(hDeviceInfoSet.ToInt32, DeviceInfoData, SPDRP_HARDWAREID, DataType, buffer, buffersize, buffersize)
                    If Marshal.GetLastWin32Error() = ERROR_INVALID_DATA Then
                        '  WriteLine("ERROR_INVALID_DATA")
                        ' May be a Legacy Device with no HardwareID. Continue.
                        Exit While

                    ElseIf Marshal.GetLastWin32Error() = ERROR_INSUFFICIENT_BUFFER Then
                        ' We need to change the buffer size.
                        ' WriteLine("ERROR_INSUFFICIENT_BUFFER")
                        ReDim buffer(0 To buffersize - 1)
                    Else
                        errore = 1
                        WriteLine("ERROR " + Marshal.GetLastWin32Error().ToString)
                        DisplayError("GetDeviceRegistryProperty", "")
                        'Unknown Failure.
                        ' MessageBox.Show("Error: " + Marshal.GetLastWin32Error().ToString, "GetDeviceRegistryProperty", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit While
                    End If
                End While

                If (Marshal.GetLastWin32Error() = ERROR_INVALID_DATA) Then
                    Continue While
                End If

                If errore = 1 Then Exit While

                If buffer.Length > 0 Then
                    Dim str As String = System.Text.Encoding.Default.GetString(buffer)
                    If str.ToUpper.IndexOf(HardwareId.ToUpper) >= 0 Then
                        WriteLine("FindExistingDevice: ok: Found: """ + HardwareId + """ at index " + i.ToString)
                        Return True
                    End If
                End If
            End While

            ' se esce per ERROR_NO_MORE_ITEMS ->   non trovato ma ok
            lastErr = Marshal.GetLastWin32Error()
            If lastErr = ERROR_NO_MORE_ITEMS Then
                WriteLine("ERROR_NO_MORE_ITEMS=" + Marshal.GetLastWin32Error.ToString)
            Else
                '  MessageBox.Show("error=" + lastErr.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                DisplayError("FindExistingDevice", "")
            End If

            SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            Return False

        Catch ex As Exception
            SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            ClassMsgErr.LogErr("ClassMdmInstall", "FindExistingDevice", ex, True)
        Finally

        End Try

    End Function

    Private Function GetDeviceRegistryProperty(ByVal DeviceInfoSet As Integer, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByRef errore As Integer) As Boolean
        Try
            Dim DataT As Integer
            Dim buffer(-1) As Byte

            Dim GetDeviceRegistryPropertyOk As Boolean = False

            Dim buffersize As Integer = 0
            'http://www.pinvoke.net/default.aspx/Constants/WINERROR.html

            While Not SetupDiGetDeviceRegistryProperty(DeviceInfoSet, DeviceInfoData, SPDRP_HARDWAREID, DataT, buffer, buffersize, buffersize)
                If Marshal.GetLastWin32Error() = ERROR_INVALID_DATA Then
                    ' May be a Legacy Device with no HardwareID. Continue.
                    Return False

                ElseIf Marshal.GetLastWin32Error() = ERROR_INSUFFICIENT_BUFFER Then
                    ' We need to change the buffer size.
                    ReDim buffer(0 To buffersize - 1)
                Else
                    errore = 1
                    'Unknown Failure.
                    MessageBox.Show("Error: " + Marshal.GetLastWin32Error().ToString, "GetDeviceRegistryProperty", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
            End While

            Return True

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "GetDeviceRegistryProperty", ex, True)
        End Try
    End Function

    Public Shared Function GetFirstFileInf() As String
        Try
            For Each f As String In My.Computer.FileSystem.GetFiles(My.Application.Info.DirectoryPath)
                If System.IO.Path.GetExtension(f).ToLower = ".inf" Then
                    Return f
                End If
            Next
            Return ""

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "GetFirstFileInf", ex, True)
        End Try

    End Function

    Private _isRunning As Boolean = False

    Public Function StartInstallDeviceAndDriver(ByVal FileInf As String, ByVal ModemPort As String) As Boolean
        Try
            If Me._isRunning = True Then
                Return False
            End If
            Dim p As New classdataTh
            p.FileInf = FileInf
            p.ModemPort = ModemPort
            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _thInstall, p)
            Return True
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "StartInstallDeviceAndDriver", ex, True)
        End Try
    End Function

    Class classdataTh
        Public FileInf As String = ""
        Public ModemPort As String = ""
    End Class

    Private Sub _thInstall(ByVal o As Object)
        Try
            Me._isRunning = True
            Dim p As classdataTh = CType(o, classdataTh)

            If Me.InstallDeviceAndDriver(p.FileInf, ClassMdmInstall.HardwareId_LS, p.ModemPort) Then
                RaiseEvent finishInstallDeviceDriver(True)
            Else
                RaiseEvent finishInstallDeviceDriver(False)
            End If
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "_thInstall", ex, True)
            RaiseEvent finishInstallDeviceDriver(False)
        Finally
            Me._isRunning = False
        End Try
    End Sub

    Public Function InstallDeviceAndDriver(ByVal FileInf As String, Hardware_ID As String, ByVal ModemPort As String) As Boolean
        Try

            Dim RebootRequired As Boolean = False
            Dim lastErr As Integer = 0

            If FileInf.Trim = "" Then
                WriteLine("Invalid name file inf")
                Return False
            End If

            If ModemPort.Trim = "" OrElse Not ModemPort.Trim.ToUpper.StartsWith("COM") Then
                WriteLine("Invalid com port")
                Return False
            End If

            If Not PortExist(ModemPort, True) Then
                WriteLine("Can't install modem: Not exist """ + ModemPort + """")
                Return False
            End If

            'HardwareId="PNPC031"
            If FindExistingDevice(Hardware_ID, lastErr) Then
                WriteLine("UpdateDriverForPlugAndPlayDevices: No Need to Create a Device Node, just update the current devnode ...")
                ' No Need to Create a Device Node, just update the current devnode.
                If Not UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, Hardware_ID, FileInf, INSTALLFLAG_READONLY, RebootRequired) Then

                    'If Not DiInstallDriver(IntPtr.Zero, FileInf, INSTALLFLAG_READONLY, RebootRequired) Then
                    DisplayError("UpdateDriverForPlugAndPlayDevices", "")
                    Return False
                Else
                    WriteLine("UpdateDriverForPlugAndPlayDevices: ok")
                    Return True
                End If
            Else

                If Not (lastErr = ERROR_NO_MORE_ITEMS) Then
                    ' An unknown failure from FindExistingDevice()
                    WriteLine("An unknown failure from FindExistingDevice(), error=" + lastErr.ToString)
                    Return False   ' Install Failure
                End If


                'Driver Does not exist, Create and call the API.


                'version 1  (originale)
                WriteLine("Not found " + Hardware_ID + " : InstallRootEnumeratedDriver ...")
                If Not InstallRootEnumeratedDriver(Hardware_ID, FileInf, ModemPort, RebootRequired) Then
                    WriteLine("InstallRootEnumeratedDriver: fail")
                    Return False  ' Install Failure
                End If


                ''version 2 
                'WriteLine("Not found " + HardwareId + " : InstallDriver ...")
                'If Not InstallDriver(HardwareId, FileInf, RebootRequired) Then
                '    WriteLine("InstallDriver: fail")
                '    Return False  ' Install Failure
                'End If


                'version 3 
                'WriteLine("Not found " + HardwareId + " : InstallModem ...")
                'If Not InstallModem(FileInf, ModemPort, HardwareId, RebootRequired) Then
                '    WriteLine("InstallModem: fail")
                '    Return False  ' Install Failure
                'End If


                Return True

            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "InstallDeviceAndDriver", ex, True)
        End Try

    End Function

    Private Function InstallRootEnumeratedDriver(ByVal HardwareId As String, ByVal FileInf As String, ByVal ModemPort As String, ByRef RebootRequired As Boolean) As Boolean
        Try
            'Modem INF File Structure:
            'https://docs.microsoft.com/en-us/previous-versions/windows/hardware/modem/ff542491%28v%3dvs.85%29


            'This routine creates and installs a new root-enumerated devnode.
            'HardwareId:  PNPC031
            'FileInf:     C : \Users\msimonato.HEMINA\AppData\Roaming\Mcp\Download\M5Driver\M5-MDM-LS.inf
            'ModemPort:   COM5
            'className:   Modem
            'ClassGUID = {4d36e96d-e325-11ce-bfc1-08002be10318}

            ' Use the INF File to extract the Class GUID. 
            Dim ClassGUID As New Guid
            Dim RequiredSize As Boolean = False
            Dim DeviceInfoData As SP_DEVINFO_DATA
            Dim className As New StringBuilder(256)
            Dim DeviceInfoSet As IntPtr
            Dim Remove As Boolean = False

            WriteLine("SetupDiGetINFClass: HardwareId=" + HardwareId + ",  FileInf= """ + FileInf + """, ModemPort=" + ModemPort + " ...")

            If SetupDiGetINFClass(FileInf, ClassGUID, className, 256, 0) = False Then
                DisplayError("SetupDiGetINFClass", "className:" + className.ToString)
                Return False
            End If

            WriteLine("SetupDiGetINFClass: ok, className=""" + className.ToString + """,  classGUID=""" + ClassGUID.ToString.ToUpper + """")

            'Create the container for the to-be-created Device Information Element.
            WriteLine("SetupDiCreateDeviceInfoList ...")
            DeviceInfoSet = SetupDiCreateDeviceInfoList(ClassGUID, IntPtr.Zero)
            If DeviceInfoSet = New IntPtr(-1) Then 'INVALID_HANDLE_VALUE 
                DisplayError("SetupDiCreateDeviceInfoList", "INVALID_HANDLE_VALUE")
                Return False
            End If

            ' Now create the element.  Use the Class GUID and Name from the INF file.
            DeviceInfoData.cbSize = Marshal.SizeOf(DeviceInfoData)

            WriteLine("SetupDiCreateDeviceInfo: DeviceInfoData.cbSize=" + DeviceInfoData.cbSize.ToString)
            '@mdmhayes.inf,%m2700%;Cavo di comunicazione tra due computer

            'If Not SetupDiCreateDeviceInfo(DeviceInfoSet, className.ToString, ClassGUID, "@mdmhayes.inf,%m2700%;Cavo di comunicazione tra due computer", IntPtr.Zero, DICD_GENERATE_ID, DeviceInfoData) Then

            If Not SetupDiCreateDeviceInfo(DeviceInfoSet, className.ToString, ClassGUID, "@mdmhayes.inf,%m2700%;Modem M5", IntPtr.Zero, DICD_GENERATE_ID, DeviceInfoData) Then
                DisplayError("SetupDiCreateDeviceInfo", "classGuid:" + DeviceInfoData.classGuid.ToString)
                SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)
                Return False
            End If

            WriteLine("DeviceInfoData.classGuid=""" + DeviceInfoData.classGuid.ToString + """")
            WriteLine("DeviceInfoData.propertyId=""" + DeviceInfoData.propertyId.ToString + """")
            WriteLine("DeviceInfoData.cbSize=""" + DeviceInfoData.cbSize.ToString + """")

            ' Do

            WriteLine("SetupDiSetDeviceRegistryProperty: " + HardwareId)

            If Not SetupDiSetDeviceRegistryProperty(DeviceInfoSet, DeviceInfoData, SPDRP_HARDWAREID, HardwareId, (HardwareId.Length + 2) * ClassMdmInstall.GetcharSize()) Then
                DisplayError("SetupDiSetDeviceRegistryProperty", "INVALID_HANDLE_VALUE")
                SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)
                Return False
            End If

            ''//////////////////////////////////////////
            'Const Fname As String = "M5"
            'If Not SetupDiSetDeviceRegistryProperty(DeviceInfoSet, DeviceInfoData, SPDRP_FRIENDLYNAME, Fname, (Fname.Length + 2) * ClassMdmInstall.GetcharSize()) Then
            '    DisplayError("SetupDiSetDeviceRegistryProperty", "INVALID_HANDLE_VALUE")
            '    Return False
            'End If
            ''//////////////////////////////////////////

            WriteLine("SetupDiSetDeviceRegistryProperty: ok")

            WriteLine("RegisterModem: ...")
            Dim drvData As New SP_DRVINFO_DATA
            If Not RegisterModem(DeviceInfoSet, DeviceInfoData, ModemPort, drvData) Then
                DisplayError("RegisterModem", "")
                Remove = True
            End If
            WriteLine("RegisterModem: ok")



            'WriteLine("DiInstallDevice: ...")
            'If Not DiInstallDevice(IntPtr.Zero, DeviceInfoSet, DeviceInfoData, drvData, 0, RebootRequired) Then
            '    DisplayError("DiInstallDevice", "")

            'End If


            WriteLine("UpdateDriverForPlugAndPlayDevices: ...")
            ' Install the Driver. INSTALLFLAG_FORCE
            If Not UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, HardwareId, FileInf, INSTALLFLAG_FORCE, RebootRequired) Then
                'If Not DiInstallDriver(IntPtr.Zero, FileInf, INSTALLFLAG_READONLY, RebootRequired) Then
                'ERROR_IN_WOW64
                If CInt(Marshal.GetLastWin32Error And &HFFFF) = ERROR_IN_WOW64 Then
                    DisplayError("UpdateDriverForPlugAndPlayDevices", "ERROR_IN_WOW64")
                Else
                    DisplayError("UpdateDriverForPlugAndPlayDevices", "")
                End If

                Remove = True
            End If
            WriteLine("UpdateDriverForPlugAndPlayDevices: ok")

            ' Loop While False



            If Remove Then
                WriteLine("Delete Device Instance")

                'Delete Device Instance that was registered using SetupDiRegisterDeviceInfo
                ' May through an error if Device not registered -- who cares??
                SetupDiCallClassInstaller(DIF_REMOVE, DeviceInfoSet, DeviceInfoData)
            End If

            ' Cleanup.
            ' Err = GetLastError()
            WriteLine("SetupDiDestroyDeviceInfoList")
            SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)
            'SetLastError(err);

            'return (err == NO_ERROR && !Remove);
            If Not Remove Then
                WriteLine("Driver and device installed succeffully!")
            End If

            Return Not Remove

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "InstallRootEnumeratedDriver", ex, True)
        End Try

    End Function

    Public Shared Function GetcharSize() As Integer
        Try

            Return Marshal.SystemDefaultCharSize
            ' Return Marshal.SizeOf(GetType(Char)).ToString
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "GetcharSize", ex, True)
        End Try
    End Function

    Public Function RegisterModem(ByVal hDeviceInfoSet As IntPtr, ByRef devInfoElem As SP_DEVINFO_DATA, ByVal PortName As String, ByRef drvData As SP_DRVINFO_DATA) As Boolean
        ' This routine registers a modem device and writes "AttachedTo" string to registry	-- this value represents the Port to which Modem is attached.
        Try
            ' hdi		-- Device Info Set -- containing the modem Device Info element.
            ' pdevData	-- Pointer the the modem device information element.
            ' szPortName-- Supplies a string containing port name the modem is to be attached to (like COM1, COM2).
            'The function returns TRUE if it was able to register the modem and write the "AttachedTo" string to registry,	otherwise it returns a FALSE.

            Dim hKeyDev As IntPtr
            Dim AttachedTo As String = "AttachedTo"
            Dim bRet As Boolean = False
            'Dim drvData As New SP_DRVINFO_DATA

            WriteLine("RegisterModem: SetupDiRegisterDeviceInfo: ...")

            If Not SetupDiRegisterDeviceInfo(hDeviceInfoSet, devInfoElem, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero) Then
                DisplayError("Register Device 1", "")
                Return False
            End If
            WriteLine("RegisterModem: SetupDiRegisterDeviceInfo: ok: " + devInfoElem.classGuid.ToString)

            WriteLine("RegisterModem: SetupDiOpenDevRegKey: ...")
            ' Opens a registry key for device-specific configuration information.
            hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ) ' This call fails....

            Dim lasterr As Integer = Marshal.GetLastWin32Error

            'if( (INVALID_HANDLE_VALUE == hKeyDev) && ( ERROR_KEY_DOES_NOT_EXIST == GetLastError()) )
            If (hKeyDev = New IntPtr(-1)) And (lasterr = ERROR_KEY_DOES_NOT_EXIST) Then
                WriteLine("RegisterModem: SetupDiOpenDevRegKey: ok")
                WriteLine("RegisterModem: SetupDiCreateDevRegKey: ...")
                hKeyDev = SetupDiCreateDevRegKey(hDeviceInfoSet, devInfoElem, DICS_FLAG_GLOBAL, 0, DIREG_DRV, IntPtr.Zero, IntPtr.Zero)

                'if( INVALID_HANDLE_VALUE == hKeyDev )
                If (hKeyDev = New IntPtr(-1)) Then
                    DisplayError("RegisterModem: SetupDiCreateDevRegKey", "")
                    Return False
                End If
                WriteLine("RegisterModem: SetupDiCreateDevRegKey: ok")
            Else
                DisplayError("RegisterModem: SetupDiOpenDevRegKey", "Can not open DriverKey")
                Return False
            End If

            Dim dwRet As Integer = 0
            Dim RegType As Microsoft.Win32.RegistryValueKind = Microsoft.Win32.RegistryValueKind.String
            Dim Size As Integer = (PortName.Length + 1) * ClassMdmInstall.GetcharSize()

            ' Dim lpdata As UInteger = CUInt(Marshal.StringToHGlobalAnsi(PortName).ToInt32)
            ' Dim ptr As IntPtr = Marshal.StringToHGlobalAuto(PortName)

            '---------------------------------------------------------------------------------------------
            ' dwRet = RegSetValueEx(hKeyDev, "FriendlyName", 0, RegType, "M5", Size)
            '---------------------------------------------------------------------------------------------


            WriteLine("RegisterModem: RegSetValueEx: name=" + AttachedTo + ",value=" + PortName + ", Size=" + Size.ToString)

            ' if (ERROR_SUCCESS != (dwRet = RegSetValueEx (hKeyDev, c_szAttachedTo, 0, REG_SZ,(PBYTE)pszPort, (lstrlen(pszPort)+1)*sizeof(TCHAR))))
            dwRet = RegSetValueEx(hKeyDev, AttachedTo, 0, RegType, PortName, Size)
            If dwRet <> ERROR_SUCCESS Then
                DisplayError("RegSetValueEx", "")
                'SetLastError (dwRet);
                bRet = False
            End If

            WriteLine("RegisterModem: RegSetValueEx: ok")

            WriteLine("RegisterModem: RegCloseKey: " + hKeyDev.ToString)

            RegCloseKey(hKeyDev)

            WriteLine("RegisterModem: SetupDiRegisterDeviceInfo: ...")
            If Not SetupDiRegisterDeviceInfo(hDeviceInfoSet, devInfoElem, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero) Then
                DisplayError("Register Device 2", "")
                Return False
            End If
            WriteLine("RegisterModem: SetupDiRegisterDeviceInfo: ok")

            WriteLine("RegisterModem: SetupDiGetSelectedDriver: ...")
            If Not SetupDiGetSelectedDriver(hDeviceInfoSet, devInfoElem, drvData) Then
                '#define ERROR_NO_DRIVER_SELECTED                 (APPLICATION_ERROR_MASK|ERROR_SEVERITY_ERROR|0x203)
                If CInt(Marshal.GetLastWin32Error And &HFFFF) = &H203 Then
                    'DisplayError("SetupDiGetSelectedDriver Failed: ", "ERROR_NO_DRIVER_SELECTED")
                    WriteLine("RegisterModem: SetupDiGetSelectedDriver: ..." + "ERROR_NO_DRIVER_SELECTED: warning")
                Else
                    DisplayError("RegisterModem: SetupDiGetSelectedDriver Failed: ", "")
                End If
            End If
            WriteLine("RegisterModem: SetupDiGetSelectedDriver: ok")

            WriteLine("RegisterModem: tutto ok")

            Return True

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "RegisterModem", ex, True)
        End Try
    End Function

    Private Function SetAtt() As Boolean
        Try
            Dim hKey As IntPtr  'unmanaged function 

        Catch ex As Exception

        End Try
    End Function

    Public Function APIErrorByNumber() As String
        Try
            Dim errore As Integer = CInt(Marshal.GetLastWin32Error And &HFFFF)
            Dim descErr As String = ""
            Dim msg As String = ""
            If CInt(errore And &HFFFF) = ERROR_IN_WOW64 Then
                descErr = "ERROR_IN_WOW64"
            End If

            msg = "Error: " + errore.ToString + " (0x" + Format(errore, "x") + ") : " + descErr + " " + New Win32Exception(errore).Message

            Return msg
        Catch ex As Exception

        End Try
    End Function

    Public Sub DisplayError(ByVal title As String, ByVal desc As String)
        Try
            Dim msg As String = ""

            'Dim errore As Integer = Marshal.GetLastWin32Error
            Dim errore As Integer = CInt(Marshal.GetLastWin32Error And &HFFFF)

            Dim descErr As String = ""

            If CInt(errore And &HFFFF) = ERROR_IN_WOW64 Then
                descErr = "ERROR_IN_WOW64"
            End If

            msg = "Error: " + errore.ToString + " (0x" + Format(errore, "x") + ") : " + vbCrLf + descErr + vbCrLf + New Win32Exception(errore).Message + vbCrLf + vbCrLf + desc

            WriteLine(msg + " " + desc)

            MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "DisplayError", ex, True)
        End Try
    End Sub

    Public Sub StartGetListModemInstalledAndFTDIPort(ByVal tree As TreeView, WaitBeforeMs As Integer, Caller As String)
        Try
            If Me._IslistingModem = False Then
                Dim DataListModemINstalled As New classDataListModemINstalled
                DataListModemINstalled.tree = tree
                DataListModemINstalled.WaitBeforeMs = WaitBeforeMs
                DataListModemINstalled.Caller = Caller

                System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _thGetListModemInstalledAndFTDIPort, DataListModemINstalled)
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub GetListModemInstalledAndFTDIPort(ByVal tree As TreeView, WaitBeforeMs As Integer, Caller As String)
        Try
            Dim DataListModemINstalled As New classDataListModemINstalled
            DataListModemINstalled.tree = tree
            DataListModemINstalled.WaitBeforeMs = WaitBeforeMs
            DataListModemINstalled.Caller = Caller

            _thGetListModemInstalledAndFTDIPort(DataListModemINstalled)
        Catch ex As Exception

        End Try
    End Sub

    Private Function FtdiBusExit() As Boolean
        Try
            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Const keyFtdibus As String = "SYSTEM\CurrentControlSet\Enum"
            For Each key_xxxx As String In rk.OpenSubKey(keyFtdibus).GetSubKeyNames
                If key_xxxx.ToUpper = "FTDIBUS" Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception

        End Try
    End Function

    Class classPortPlugged
        Public Name As String = ""
        Public Plugged As Boolean = False
        Public FriendlyName As String = ""
        Public VID As String = ""
        Public PID As String = ""
        Public Driver As String = ""

        Public Sub New(ByVal Name As String, ByVal Plugged As Boolean, FriendlyName As String, VID As String, PID As String, Driver As String)
            Me.Name = Name
            Me.Plugged = Plugged
            Me.FriendlyName = FriendlyName
            Me.VID = VID
            Me.PID = PID
            Me.Driver = Driver
        End Sub

    End Class



    ''' <summary>
    ''' Get USB FTDI chip PID 6015,7018,7019 plugged.
    ''' </summary>
    ''' <param name="OnlyWithM5_speedReg">If exist register M5_Speed in Device parameters, add device to list</param>
    ''' <returns>List of device connected</returns>
    Public Function GetFTDIPlugged(OnlyWithM5_speedReg As Boolean) As List(Of classPortPlugged)
        Try
            Dim dtms As String = ""
            Return GetFTDIPlugged(OnlyWithM5_speedReg, dtms)
        Catch ex As Exception

        End Try
    End Function

    ''' <summary>
    ''' Get USB FTDI chip PID 6015,7018,7019 plugged.
    ''' </summary>
    ''' <param name="OnlyWithM5_speedReg">If exist register M5_Speed in Device parameters, add device to list</param>
    ''' <param name="dtms">Time to read ms</param>
    ''' <returns>List of device connected</returns>
    Public Function GetFTDIPlugged(OnlyWithM5_speedReg As Boolean, ByRef dtms As String) As List(Of classPortPlugged)

        Dim listPortFTDI_6015Installed As New List(Of classPortPlugged)
        Dim _sw As New Stopwatch
        Try

            _sw.Reset()
            _sw.Start()

            Dim n As Integer = 0

            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            n = 0

            If FtdiBusExit() Then

                ' port
                For Each key_xxxx As String In rk.OpenSubKey(keyFtdibus).GetSubKeyNames

                    If key_xxxx.ToUpper.StartsWith(ClassFtdiRegKey.FtdiVidPid7018) Or
                       key_xxxx.ToUpper.StartsWith(ClassFtdiRegKey.FtdiVidPid7019) Or
                       key_xxxx.ToUpper.StartsWith(ClassFtdiRegKey.FtdiVidPid6015) Then

                        For Each keyS As String In rk.OpenSubKey(keyFtdibus + "\" + key_xxxx).GetSubKeyNames
                            Try
                                Dim PID As String = ""
                                Dim VID As String = ""
                                ClassFtdiRegKey.EstraiVidPid(key_xxxx.ToUpper, ClassFtdiRegKey.EnumVidPidSeparator.Più, VID, PID)
                                n += 1
                                Dim FriendlyName As String = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS).GetValue("FriendlyName").ToString
                                Dim Driver As String = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS).GetValue("Driver").ToString

                                Dim PortNameValue As String = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValue("PortName").ToString

                                Dim ns() As String = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValueNames

                                If OnlyWithM5_speedReg Then
                                    Dim M5Speed As String = ""
                                    If ClassFtdiRegKey.ReadM5_Speed(ns, rk, keyFtdibus, key_xxxx, keyS, M5Speed) Then
                                        If Me.PortExist(PortNameValue, False) Then
                                            listPortFTDI_6015Installed.Add(New classPortPlugged(PortNameValue, True, FriendlyName, VID, PID, Driver))
                                        Else
                                            listPortFTDI_6015Installed.Add(New classPortPlugged(PortNameValue, False, FriendlyName, VID, PID, Driver))
                                        End If

                                    End If
                                Else
                                    If Me.PortExist(PortNameValue, False) Then
                                        listPortFTDI_6015Installed.Add(New classPortPlugged(PortNameValue, True, FriendlyName, VID, PID, Driver))
                                    End If
                                End If







                            Catch ex As Exception

                            End Try

                        Next

                    End If
                Next
            End If

        Catch ex As Exception
        Finally
            _sw.Stop()
            dtms = Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms"
        End Try

        Return listPortFTDI_6015Installed

    End Function

    Private Sub AddFTDIToTree(ByVal tree As TreeView, ByVal treePort As TreeNode)
        Try

            If tree IsNot Nothing Then



                For i As Integer = 0 To Me.listPortFTDI.Count - 1

                    Dim p As classPortFTDI
                    p = Me.listPortFTDI(i)

                    If p.Usb IsNot Nothing Then

                        Dim nodePadre As New Safe.TreeView.ClassTreeData
                        Safe.TreeView.NodesAddImg(p.COM, "(" + (i + 1).ToString + "): " + p.COM + " [" + p.DescPid + "]", treePort, 2, nodePadre)
                        Safe.TreeView.NodeTag(p.COM, nodePadre.Nodo)

                        ' Is port Plugged
                        If p.IsPlugged Then

                            Safe.TreeView.TreeNodesClear(tree, nodePadre.Nodo)
                            Dim nodePr As New Safe.TreeView.ClassTreeData
                            Dim spd As String = p.Usb.M5Speed
                            If p.Usb.M5Speed = "" Then
                                spd = ""
                            Else
                                spd = ", " + p.Usb.M5Speed
                            End If
                            Safe.TreeView.NodesAddImg(p.Usb.USBdescriprion, p.Usb.USBraw.Replace("HEMINA-", "") + " (PID:" + p.Usb.PID + spd + ")", nodePadre.Nodo, 5, nodePr)
                            Safe.TreeView.NodeTag(p.Usb.USBdescriprion, nodePr.Nodo)


                            If p.PID = "6015" Then
                                nodePr.Nodo.BackColor = Color.Yellow
                                nodePr.Nodo.ForeColor = Color.Blue
                            Else
                                nodePr.Nodo.BackColor = Color.Azure
                                nodePr.Nodo.ForeColor = Color.Blue
                            End If


                        Else
                            'nodePr.Nodo.BackColor = Color.White
                            'nodePr.Nodo.ForeColor = Color.Black

                        End If
                    Else
                        'Dim nodePadre As New Safe.TreeView.ClassTreeData
                        'Safe.TreeView.NodesAddImg(p.COM, "(" + (i + 1).ToString + "): " + p.COM + " [" + p.DescPid + "]", treePort, 2, nodePadre)
                        'Safe.TreeView.NodeTag(p.COM, nodePadre.Nodo)

                        'If p.IsPlugged Then
                        '    If p.PID = "6015" Then
                        '        'nodePadre.Nodo.BackColor = Color.Yellow
                        '        nodePadre.Nodo.ForeColor = Color.Red
                        '    Else
                        '        ' nodePadre.Nodo.BackColor = Color.Azure
                        '        nodePadre.Nodo.ForeColor = Color.Blue
                        '    End If
                        'End If



                        ' Dim portNode As TreeNode = treePort.Nodes.Add("", "(" + n.ToString + "): " + devPar, 2, 2)
                    End If

                Next

            End If




        Catch ex As Exception
            MessageBox.Show(ex.Message, "AddFTDIToTree", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Function AddPortNameFTDIInstalledAndPlugged(key_xxxx As String, DescPid As String, VID_PID As String, rk As Microsoft.Win32.RegistryKey, ByRef tree As TreeView, ByRef treePort As TreeNode, ByRef n As Integer) As Boolean
        Try
            For Each keyS As String In rk.OpenSubKey(keyFtdibus + "\" + key_xxxx).GetSubKeyNames

                Try

                    n += 1

                    Dim PortNameValue As String = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValue("PortName").ToString

                    Dim M5Speed As String = ""

                    Dim ns(-1) As String

                    ns = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValueNames

                    If ClassFtdiRegKey.ReadM5_Speed(ns, rk, keyFtdibus, key_xxxx, keyS, M5Speed) Then
                    End If

                    Dim p As New classPortFTDI
                    p.COM = PortNameValue
                    p.PID = VID_PID
                    p.DescPid = DescPid

                    Me.listPortFTDI.Add(p)

                    If tree IsNot Nothing Then

                        Dim nodePr As New Safe.TreeView.ClassTreeData
                        Safe.TreeView.NodesAddImg(PortNameValue, "(" + n.ToString + "): " + PortNameValue + " [" + DescPid + "]", treePort, 2, nodePr)
                        Safe.TreeView.NodeTag(PortNameValue, nodePr.Nodo)

                        ' Is port Plugged
                        If Me.PortExist(PortNameValue, False) Then
                            nodePr.Nodo.BackColor = Color.Azure
                            nodePr.Nodo.ForeColor = Color.Blue
                            p.IsPlugged = True
                        Else
                            nodePr.Nodo.BackColor = Color.White
                            nodePr.Nodo.ForeColor = Color.Black
                            p.IsPlugged = False
                        End If

                        ' Dim portNode As TreeNode = treePort.Nodes.Add("", "(" + n.ToString + "): " + devPar, 2, 2)

                    End If



                Catch ex As Exception

                End Try

            Next
        Catch ex As Exception

        End Try
    End Function

    Private Function AddPortNameFTDIInstalledAndPlugged(key_xxxx As String, DescPid As String, VID_PID As String, rk As Microsoft.Win32.RegistryKey, ByRef n As Integer) As Boolean
        Try
            For Each keyS As String In rk.OpenSubKey(keyFtdibus + "\" + key_xxxx).GetSubKeyNames

                Try

                    n += 1

                    Dim PortNameValue As String = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValue("PortName").ToString

                    Dim M5Speed As String = ""

                    Dim ns(-1) As String

                    ns = rk.OpenSubKey(keyFtdibus + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValueNames

                    If ClassFtdiRegKey.ReadM5_Speed(ns, rk, keyFtdibus, key_xxxx, keyS, M5Speed) Then
                    End If

                    Dim p As New classPortFTDI
                    p.COM = PortNameValue
                    p.PID = VID_PID
                    p.DescPid = DescPid

                    Me.listPortFTDI.Add(p)


                    ' Is port Plugged
                    If Me.PortExist(PortNameValue, False) Then
                        p.IsPlugged = True
                    Else
                        p.IsPlugged = False
                    End If


                Catch ex As Exception

                End Try

            Next
        Catch ex As Exception

        End Try
    End Function

    Class classDataListModemINstalled
        Public tree As TreeView = Nothing
        Public WaitBeforeMs As Integer = 0
        Public Caller As String = ""
    End Class

    Private Sub _thGetListModemInstalledAndFTDIPort(ByVal o As Object)


        Dim DataListModemINstalled As classDataListModemINstalled = Nothing
        Try
            DataListModemINstalled = CType(o, classDataListModemINstalled)

            Me._IslistingModem = True

            Safe.Control.Visible(False, _btnRefresh)
            Safe.Control.Visible(False, Me._tablLytPanelUnistall)


            DataListModemINstalled = CType(o, classDataListModemINstalled)

            Const keyModem As String = "SYSTEM\CurrentControlSet\Control\Class\{4D36E96D-E325-11CE-BFC1-08002BE10318}"
            Const keyAttachedTo As String = "AttachedTo"
            Const keyFriendlyName As String = "FriendlyName"
            Const keyDCB As String = "DCB"
            Const keyDriverVersion As String = "DriverVersion"

            Safe.TreeView.NodesClear(DataListModemINstalled.tree)

            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim numero As Integer = 0
            Dim n As Integer = 0
            Dim treePc As TreeNode = Nothing
            Dim treeModem As TreeNode = Nothing
            Dim treePort As TreeNode = Nothing
            Dim nodePc As New Safe.TreeView.ClassTreeData
            Dim nodeModem As New Safe.TreeView.ClassTreeData
            Dim nodePort As New Safe.TreeView.ClassTreeData

            If DataListModemINstalled.tree IsNot Nothing Then
                Safe.TreeView.TreeAddImg(My.Computer.Name, DataListModemINstalled.tree, 0, nodePc)
                treePc = nodePc.Nodo

                Safe.TreeView.NodesAddImg("", "Port M5:", treePc, 4, nodePort)
                treePort = nodePort.Nodo

                Safe.TreeView.NodesAddImg("", "Modem M5:", treePc, 1, nodeModem)
                treeModem = nodeModem.Nodo

                Safe.TreeView.ExpandAll(DataListModemINstalled.tree)

            End If

            ' wait before read
            System.Threading.Thread.Sleep(DataListModemINstalled.WaitBeforeMs)

            n = 0
            Me.listPortFTDI.Clear()


            If FtdiBusExit() Then

                Dim sbn() As String = rk.OpenSubKey(keyFtdibus).GetSubKeyNames
                'SYSTEM\CurrentControlSet\Enum\FTDIBUS
                ' PortName
                For Each key_xxxx As String In rk.OpenSubKey(keyFtdibus).GetSubKeyNames

                    If key_xxxx.ToUpper.StartsWith(ClassFtdiRegKey.FtdiVidPid7018) Then
                        ' AddPortNameFTDIInstalledAndPlugged(key_xxxx, "LS", "7018", rk, DataListModemINstalled.tree, treePort, n)
                        AddPortNameFTDIInstalledAndPlugged(key_xxxx, "LS", "7018", rk, n)
                    End If

                    If key_xxxx.ToUpper.StartsWith(ClassFtdiRegKey.FtdiVidPid7019) Then
                        ' AddPortNameFTDIInstalledAndPlugged(key_xxxx, "HS", "7019", rk, DataListModemINstalled.tree, treePort, n)
                        AddPortNameFTDIInstalledAndPlugged(key_xxxx, "HS", "7019", rk, n)
                    End If

                    If key_xxxx.ToUpper.StartsWith(ClassFtdiRegKey.FtdiVidPid6015) Then
                        ' AddPortNameFTDIInstalledAndPlugged(key_xxxx, "6015", "6015", rk, DataListModemINstalled.tree, treePort, n)
                        AddPortNameFTDIInstalledAndPlugged(key_xxxx, "6015", "6015", rk, n)
                    End If
                Next

                Me.WriteLine("Found " + listPortFTDI.Count.ToString + " port installed, " + classPortFTDI.GetNumPortsPlugged(Me.listPortFTDI).ToString + " port plugged")
            Else
                Me.WriteLine("M5 driver not found.")
            End If



            If classPortFTDI.GetNumPortsPlugged(Me.listPortFTDI) = 0 Then
                '  TopMostMessageBox.ShowNoBlock("No usb port plugged!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                'System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _thGetFTDIinfo_USBRaw, tree)
                Safe.TreeView.ExpandAll(DataListModemINstalled.tree)
            End If


            ' Read USB Raw string from driver
            _thGetFTDIinfo_USBRaw(DataListModemINstalled.tree)


            n = 0
            Me.listMdmInstalledName.Clear()

            ' modem
            For Each key_xxxx As String In rk.OpenSubKey(keyModem).GetSubKeyNames
                If key_xxxx IsNot Nothing Then
                    If Integer.TryParse(key_xxxx, numero) Then
                        n += 1
                        Dim AttachedTo As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyAttachedTo).ToString
                        If AttachedTo IsNot Nothing Then
                            Dim FriendlyName As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyFriendlyName).ToString
                            Dim DCB() As Byte = CType(rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyDCB), Byte())
                            Dim BaudRate As String = DCBtoBaudRate(DCB)
                            Dim DriverVersion As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyDriverVersion).ToString

                            If FriendlyName IsNot Nothing Then
                                'If IsInList(AttachedTo, Me.listPortFTDI_Installed) Then
                                Dim mdm As New LibRasBook.ClassRasBook.classMdmInstalled
                                mdm.Name = FriendlyName
                                mdm.ComPort = AttachedTo
                                Me.listMdmInstalledName.Add(mdm)
                                If DataListModemINstalled.tree IsNot Nothing Then
                                    Dim nodeMdm As New Safe.TreeView.ClassTreeData
                                    Safe.TreeView.NodesAddImg(AttachedTo, "(" + n.ToString + ") " + AttachedTo + ": " + FriendlyName + ", [" + BaudRate + "] Vers.:" + DriverVersion, treeModem, 3, nodeMdm)
                                    Safe.TreeView.NodeTag(AttachedTo, nodeMdm.Nodo)
                                End If
                                ' End If
                            End If
                        End If
                    End If
                End If
            Next

            Me.WriteLine("Found " + listMdmInstalledName.Count.ToString + " modem installed")

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall -> " + DataListModemINstalled.Caller, "_thGetListModemInstalledAndFTDIPort", ex, True)
        Finally
            Safe.TreeView.ExpandAll(DataListModemINstalled.tree)
            Safe.Control.Visible(True, _btnRefresh)

            RaiseEvent TreeviewPortUpdated(Me.listPortFTDI.Count)
            Me._IslistingModem = False
        End Try

    End Sub
    Private Function IsInUsbList(COM As String, USB As ClassSerachHeminaUSBraw.classUSB) As Boolean
        Try
            For i As Integer = 0 To USB.DeviceList.Count - 1
                If USB.DeviceList(i).COM.ToUpper = COM.ToUpper Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception

        End Try
    End Function

    Private Function DCBtoBaudRate(DCB() As Byte) As String
        Try
            If DCB IsNot Nothing Then

                Dim b(0 To 1) As Byte
                b(0) = DCB(4)
                b(1) = DCB(5)
                Return BitConverter.ToUInt16(b, 0).ToString
            Else
                Return ""

            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private _lockObj As New Object
    Private USB As New ClassSerachHeminaUSBraw.classUSB

    Private Sub _thGetFTDIinfo_USBRaw(o As Object)
        Dim tree As TreeView = Nothing
        Try
            tree = CType(o, TreeView)

            SyncLock Me._lockObj

                Safe.Control.Visible(True, _lblReadingRawString)

                Dim SerachHeminaUSBraw As New ClassSerachHeminaUSBraw

                If SerachHeminaUSBraw.SearchHeminaUSB(SerachHeminaUSBraw.FullPath_M5UsbListexe, SerachHeminaUSBraw.NomeFileToSave, USB) Then


                    'Dim ftdi As New LibM5Ftdi.ClassFTDI
                    'Dim l As List(Of LibM5Ftdi.ClassFTDI.ClassUsbM5)
                    'Dim dtms As String = ""

                    'Dim _sw As New Stopwatch
                    '_sw.Reset()
                    '_sw.Start()

                    'Do
                    '    l = ftdi.SearchM5(dtms)
                    '    If l.Count > 0 Then
                    '        Exit Do
                    '    End If
                    '    If Me.listPortFTDI_6015Plugged.Count = 0 Then
                    '        Exit Do
                    '    End If
                    '    If _sw.Elapsed.TotalMilliseconds >= 10000 Then
                    '        Exit Do
                    '    End If
                    '    System.Threading.Thread.Sleep(400)
                    'Loop


                    'For i As Integer = 0 To Me.listPortFTDI_6015Plugged.Count - 1

                    'Next

                    Dim lhemina As New List(Of classPortFTDI)

                    For i As Integer = 0 To Me.listPortFTDI.Count - 1
                        For j As Integer = 0 To USB.DeviceList.Count - 1
                            If Me.listPortFTDI(i).COM.Trim = USB.DeviceList(j).COM.Trim Then
                                Me.listPortFTDI(i).Usb = CType(USB.DeviceList(j).Clone, ClassSerachHeminaUSBraw.classUSBraw)
                            End If
                        Next
                    Next

                    AddFTDIToTree(tree, tree.Nodes(0).Nodes(0))

                    'For i As Integer = 0 To USB.DeviceList.Count - 1

                    '    For j As Integer = 0 To tree.Nodes(0).Nodes(0).Nodes.Count - 1

                    '        Dim nodePadre As TreeNode = tree.Nodes(0).Nodes(0).Nodes(j)

                    '        If nodePadre.Name.Trim = USB.DeviceList(i).COM.Trim Then
                    '            Safe.TreeView.TreeNodesClear(tree, nodePadre)
                    '            Dim nodePr As New Safe.TreeView.ClassTreeData
                    '            Safe.TreeView.NodesAddImg(USB.DeviceList(i).USBdescriprion, USB.DeviceList(i).USBraw.Replace("HEMINA-", "") + " (PID:" + USB.DeviceList(i).PID + ", " + USB.DeviceList(i).M5Speed + ")", nodePadre, 5, nodePr)
                    '            Safe.TreeView.NodeTag(USB.DeviceList(i).USBdescriprion, nodePr.Nodo)
                    '            If USB.DeviceList(i).PID = "6015" Then
                    '                nodePadre.BackColor = Color.Yellow
                    '                nodePadre.ForeColor = Color.Blue
                    '            Else
                    '                nodePadre.BackColor = Color.Azure
                    '                nodePadre.ForeColor = Color.Blue
                    '            End If

                    '        End If
                    '    Next
                    'Next
                End If

                Safe.Control.Visible(False, _lblReadingRawString)

                ' stampa contenuto file usb list
                Safe.Control.Text("Usb deivices: " + vbCrLf + vbCrLf + USB.FileContent + vbCrLf + vbCrLf + USB.FilePath, Me._txtUsbRaw)

                ' cerca device con PID 6015 e propone di cambiarlo in 7018 e 7019
                Dim l As New List(Of ClassSerachHeminaUSBraw.classUSBraw)
                Dim s As String = "Found:" + vbCrLf + vbCrLf
                For i As Integer = 0 To USB.DeviceList.Count - 1
                    If USB.DeviceList(i).PID = "6015" Then
                        s = s + USB.DeviceList(i).USBraw + " with USB PID: " + USB.DeviceList(i).PID + vbCrLf
                        l.Add(USB.DeviceList(i))
                    End If
                Next

                If l.Count > 0 Then
                    s = s + vbCrLf + "Would you like to change USB PID to 7018 or 7019?"
                    If TopMostMessageBox.Show(s, "Attention!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                        Dim ftdi As New LibM5Ftdi.ClassFTDI
                        Dim listUsbFtdi As List(Of LibM5Ftdi.ClassFTDI.ClassUsbM5)
                        Dim dtms As String = ""
                        listUsbFtdi = ftdi.ChangePID(dtms)

                        If listUsbFtdi.Count > 0 Then
                            Dim desc As String = ""
                            For j As Integer = 0 To listUsbFtdi.Count - 1
                                desc = desc + listUsbFtdi(j).Description + ": PID: " + listUsbFtdi(j).PID + vbCrLf
                            Next
                            TopMostMessageBox.ShowNoBlock(desc + vbCrLf + vbCrLf + "Unplug and replug USB.", "PID changed!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            ' c'è stato un errore
                            '  TopMostMessageBox.ShowNoBlock("Error changing PID." + vbCrLf + vbCrLf + "Unplug and replug USB.", "PID change", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                    End If
                End If

            End SyncLock

        Catch ex As Exception

        Finally
            Safe.Control.Visible(False, _lblReadingRawString)
        End Try

    End Sub

    Private Sub _thGetFTDIinfoFTDI(o As Object)
        Dim tree As TreeView = Nothing
        Try
            tree = CType(o, TreeView)

            SyncLock Me._lockObj



                Dim ftdi As New LibM5Ftdi.ClassFTDI
                Dim l As List(Of LibM5Ftdi.ClassFTDI.ClassUsbM5)
                Dim dtms As String = ""

                Dim _sw As New Stopwatch
                _sw.Reset()
                _sw.Start()

                Do
                    l = ftdi.SearchM5(dtms)
                    If l.Count > 0 Then
                        Exit Do
                    End If
                    If classPortFTDI.GetNumPortsPlugged(Me.listPortFTDI) = 0 Then
                        Exit Do
                    End If
                    If _sw.Elapsed.TotalMilliseconds >= 10000 Then
                        Exit Do
                    End If
                    System.Threading.Thread.Sleep(400)
                Loop


                'For i As Integer = 0 To Me.listPortFTDI_6015Plugged.Count - 1

                'Next

                For i As Integer = 0 To l.Count - 1
                    For j As Integer = 0 To tree.Nodes(0).Nodes(0).Nodes.Count - 1

                        Dim nodePadre As TreeNode = tree.Nodes(0).Nodes(0).Nodes(j)

                        If nodePadre.Name = l(i).ComPort Then
                            Safe.TreeView.TreeNodesClear(tree, nodePadre)
                            Dim nodePr As New Safe.TreeView.ClassTreeData
                            Safe.TreeView.NodesAddImg(l(i).ID.ToString, l(i).Description.Replace("HEMINA-", ""), nodePadre, 5, nodePr)
                            Safe.TreeView.NodeTag(l(i).ID.ToString, nodePr.Nodo)
                        End If
                    Next
                Next



            End SyncLock
        Catch ex As Exception

        End Try
    End Sub

    Private Function IsInList(ByVal val As String, ByVal l As List(Of String)) As Boolean
        Try
            For i As Integer = 0 To l.Count - 1
                If l(i).ToUpper = val.ToUpper Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception

        End Try
    End Function

    Public Shared Function Get32or64bit() As String
        Try
            Return (IntPtr.Size * 8).ToString + " bit"
        Catch ex As Exception

        End Try
    End Function

    Public Sub New(ByVal txtLog As TextBox, PathLog As String, lblReadingRawString As Label, txtUsbRaw As TextBox, btnRefresh As Button, tablLytPanelUnistall As TableLayoutPanel)
        Me.txtLog = txtLog
        Me._lblReadingRawString = lblReadingRawString
        Me._txtUsbRaw = txtUsbRaw
        Me._btnRefresh = btnRefresh
        Me._tablLytPanelUnistall = tablLytPanelUnistall
        Me.PathLog = PathLog
        Me.ResetLog()
    End Sub
End Class
