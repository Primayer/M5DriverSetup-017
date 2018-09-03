
Imports System.Text
Imports System.Runtime.InteropServices



Partial Public Class ClassMdmInstall

    'https://groups.google.com/forum/#!topic/microsoft.public.development.device.drivers/W4awiLUP6F0

    'Windows API Reference: Functions:
    'http://www.jasinskionline.com/windowsapi/ref/funca.html

    Public Const DI_FLAGSEX_ALLOWEXCLUDEDDRVS As Integer = &H800
    Public Const DI_FLAGSEX_ALWAYSWRITEIDS As Integer = &H200
    Public DI_ENUMSINGLEINF As Integer = &H10000
    ' flags for device installation
    Public Const DI_QUIETINSTALL As Integer = &H800000     ' don't confuse the user with questions Or excess info


    Public Function InstallModem(ByVal INF_File As String, ByVal COM_Port As String, ByVal Hardware_ID As String, ByRef RebootRequired As Boolean) As Boolean
        Const descF As String = "InstallModem: "
        Try
            'http://www.sql.ru/forum/632921/vozmozhno-li-programmno-ustanovit-modem

            Dim bResult As Boolean = False
            Dim m_ClassGUID As Guid
            Dim m_ClassName As New StringBuilder(256)
            Dim ReqSize As UInteger = 0
            Dim hDeviceInfoSet As IntPtr
            Dim m_DeviceInfoData As New SP_DEVINFO_DATA
            Dim bRemove As Boolean = False
            Dim hKeyDev As IntPtr
            Dim m_DriverInfoData As New SP_DRVINFO_DATA
            Dim m_DeviceInstallParams As New SP_DEVINSTALL_PARAMS

            ' Use the INF File to extract the Class GUID
            WriteLine(descF + "SetupDiGetINFClass: HardwareId=" + Hardware_ID + ",  FileInf= """ + INF_File + """ Com port:" + COM_Port + " ...")


            If SetupDiGetINFClass(INF_File, m_ClassGUID, m_ClassName, 256, 0) = False Then
                DisplayError("SetupDiGetINFClass", "className:" + m_ClassName.ToString)
                Return False
            End If

            'bResult = SetupDiGetINFClass(INF_File, m_ClassGUID, m_ClassName, 256, ReqSize)

            'If (bResult = False) And (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
            '    m_ClassName = New StringBuilder(CInt(ReqSize))
            '    bResult = SetupDiGetINFClass(INF_File, m_ClassGUID, m_ClassName, CInt(ReqSize))
            '    If bResult = False Then
            '        WriteLine("SetupDiGetINFClass: " & APIErrorByNumber())
            '        Return False
            '    Else
            '        'm_ClassName = Left(m_ClassName.ToString, InStr(m_ClassName.ToString, Chr(0)) - 1)
            '        WriteLine("m_ClassName=" & m_ClassName.ToString)
            '    End If
            'Else
            '    WriteLine("SetupDiGetINFClass: " & APIErrorByNumber())
            '    Return False
            'End If

            WriteLine(descF + " ok, className=""" + m_ClassName.ToString + """,  classGUID=""" + m_ClassGUID.ToString.ToUpper + """")

            'Create the container for the to-be-created Device Information Element.
            WriteLine(descF + "SetupDiCreateDeviceInfoList ...")
            hDeviceInfoSet = SetupDiCreateDeviceInfoList(m_ClassGUID, IntPtr.Zero)
            If hDeviceInfoSet = New IntPtr(-1) Then 'INVALID_HANDLE_VALUE 
                DisplayError("SetupDiCreateDeviceInfoList", "INVALID_HANDLE_VALUE")
                Return False
            End If

            ' Now create the element. Use the Class GUID and Name from the INF file.
            m_DeviceInfoData.cbSize = Marshal.SizeOf(m_DeviceInfoData)

            WriteLine(descF + "DeviceInfoData.cbSize=" + m_DeviceInfoData.cbSize.ToString + " ...")
            If Not SetupDiCreateDeviceInfo(hDeviceInfoSet, m_ClassName.ToString, m_ClassGUID, "@mdmhayes.inf,%m2700%;Modem M5", IntPtr.Zero, DICD_GENERATE_ID, m_DeviceInfoData) Then
                DisplayError("SetupDiCreateDeviceInfo", "classGuid:" + m_DeviceInfoData.classGuid.ToString)
                GoTo Cleanup
            End If


            ' Add the HardwareID to the Device's HardwareID property.
            WriteLine(descF + "DeviceInfoData.classGuid=""" + m_DeviceInfoData.classGuid.ToString + """")
            WriteLine(descF + "DeviceInfoData.propertyId=""" + m_DeviceInfoData.propertyId.ToString + """")
            WriteLine(descF + "DeviceInfoData.cbSize=""" + m_DeviceInfoData.cbSize.ToString + """")

            WriteLine(descF + "SetupDiSetDeviceRegistryProperty: " + Hardware_ID)
            If Not SetupDiSetDeviceRegistryProperty(hDeviceInfoSet, m_DeviceInfoData, SPDRP_HARDWAREID, Hardware_ID, (Hardware_ID.Length + 2) * ClassMdmInstall.GetcharSize()) Then
                DisplayError("SetupDiSetDeviceRegistryProperty", "INVALID_HANDLE_VALUE")
                GoTo Cleanup
            End If

            WriteLine(descF + "SetupDiRegisterDeviceInfo: ...")

            If Not SetupDiRegisterDeviceInfo(hDeviceInfoSet, m_DeviceInfoData, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero) Then
                DisplayError("Register Device 1", "")
                bRemove = True
                GoTo Cleanup
            End If
            WriteLine(descF + "SetupDiRegisterDeviceInfo: ok: " + m_DeviceInfoData.classGuid.ToString)


            WriteLine(descF + "SetupDiOpenDevRegKey ...")
            ' Opens a registry key for device-specific configuration information.
            '  hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, m_DeviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ)

            hKeyDev = SetupDiOpenDevRegKey(hDeviceInfoSet, m_DeviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_ALL_ACCESS)

            Dim lasterr As Integer = Marshal.GetLastWin32Error

            If (hKeyDev = New IntPtr(-1)) And (lasterr = ERROR_KEY_DOES_NOT_EXIST) Then 'This call fails....
                WriteLine(descF + "SetupDiOpenDevRegKey: ok")
                WriteLine(descF + "SetupDiCreateDevRegKey ...")
                hKeyDev = SetupDiCreateDevRegKey(hDeviceInfoSet, m_DeviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DRV, IntPtr.Zero, IntPtr.Zero)

                'if( INVALID_HANDLE_VALUE == hKeyDev )
                If (hKeyDev = New IntPtr(-1)) Then
                    DisplayError(descF + "SetupDiCreateDevRegKey", "")
                    bRemove = True
                    GoTo Cleanup
                End If
                WriteLine(descF + "SetupDiCreateDevRegKey: ok")
            Else
                DisplayError(descF + "SetupDiOpenDevRegKey", "Can not open DriverKey")
                bRemove = True
                GoTo Cleanup
            End If

            Dim dwRet As Integer = 0
            Dim RegType As Microsoft.Win32.RegistryValueKind = Microsoft.Win32.RegistryValueKind.String
            Dim Size As Integer = (COM_Port.Length + 1) * ClassMdmInstall.GetcharSize()
            Dim AttachedTo As String = "AttachedTo"
            Dim bRet As Boolean = False
            WriteLine(descF + "RegSetValueEx: name=" + AttachedTo + ",value=" + COM_Port + ", Size=" + Size.ToString)

            ' if (ERROR_SUCCESS != (dwRet = RegSetValueEx (hKeyDev, c_szAttachedTo, 0, REG_SZ,(PBYTE)pszPort, (lstrlen(pszPort)+1)*sizeof(TCHAR))))
            dwRet = RegSetValueEx(hKeyDev, AttachedTo, 0, RegType, COM_Port, Size)
            If dwRet <> ERROR_SUCCESS Then
                DisplayError("RegSetValueEx", "")
                'SetLastError (dwRet);
                bRemove = True
                GoTo Cleanup
            End If

            WriteLine(descF + "RegSetValueEx: ok")

            WriteLine(descF + "RegCloseKey: " + hKeyDev.ToString)

            RegCloseKey(hKeyDev)

            WriteLine(descF + "SetupDiRegisterDeviceInfo ...")
            If Not SetupDiRegisterDeviceInfo(hDeviceInfoSet, m_DeviceInfoData, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero) Then
                DisplayError("Register Device 2", "")
                bRemove = True
                GoTo Cleanup
            End If
            WriteLine(descF + "SetupDiRegisterDeviceInfo: ok")

            ' Install the Driver: primo tentativo (InstallSelectedDriver) -installa il driver solo per questo dispositivo
            '  m_DeviceInstallParams installa il driver di un Inf strettamente specificato

            ' InstallSelectedDriver works on the selected device and on the
            ' selected driver on that device. Therefore, set this device as the selected one in the device information list.
            WriteLine(descF + "SetupDiSetSelectedDevice ...")
            bResult = SetupDiSetSelectedDevice(hDeviceInfoSet, m_DeviceInfoData)
            If bResult = False Then
                WriteLine("SetupDiSetSelectedDevice: " & APIErrorByNumber())
                GoTo ForceInstall
            End If
            WriteLine(descF + "SetupDiSetSelectedDevice: ok")

            '======================================================================================================================================================
            ' You now have a SP_DEVINFO_DATA structure representing your device.  Next, get a SP_DRVINFO_DATA structure to install on that device.
            m_DeviceInstallParams = New SP_DEVINSTALL_PARAMS
            m_DeviceInstallParams.cbSize = Marshal.SizeOf(m_DeviceInstallParams)
            'ReDim m_DeviceInstallParams.DriverPath(0 To MAX__DESC - 1)


            WriteLine(descF + "SetupDiGetDeviceInstallParams ...")
            bResult = SetupDiGetDeviceInstallParams(hDeviceInfoSet, m_DeviceInfoData, m_DeviceInstallParams)
            If bResult = False Then
                WriteLine("SetupDiGetDeviceInstallParams: " & APIErrorByNumber())
                '  GoTo ForceInstall
            End If

            ' Only build the driver list out of the passed-in INF.
            ' To do this, set the DI_ENUMSINGLEINF flag, and copy the
            ' full path of the INF into the DriverPath field of the
            ' DeviceInstallParams structure.
            m_DeviceInstallParams.Flags = m_DeviceInstallParams.Flags Or DI_ENUMSINGLEINF
            m_DeviceInstallParams.DriverPath = INF_File
            'ReDim m_DeviceInstallParams.DriverPath(0 To MAX__DESC - 1)
            'For i As Integer = 0 To Len(INF_File) - 1
            '    m_DeviceInstallParams.DriverPath(i) = CByte(Asc(Mid(INF_File, i + 1, 1)))
            'Next i


            ' Set the DI_FLAGSEX_ALLOWEXCLUDEDDRVS flag so that you can use this INF even if it is marked as ExcludeFromSelect.
            ' ExcludeFromSelect means do not show the INF in the legacy Add Hardware Wizard.
            m_DeviceInstallParams.FlagsEx = m_DeviceInstallParams.FlagsEx Or DI_FLAGSEX_ALLOWEXCLUDEDDRVS
            WriteLine(descF + "SetupDiSetDeviceInstallParams ...")
            bResult = SetupDiSetDeviceInstallParams(hDeviceInfoSet, m_DeviceInfoData, m_DeviceInstallParams)
            If bResult = False Then
                WriteLine("SetupDiSetDeviceInstallParams: " & APIErrorByNumber())
                ' GoTo ForceInstall
            End If
            '==================================================

            ' Build up a Driver Information List.
            ' Build a compatible driver list, meaning only include the driver nodes that match one of the hardware or compatible Ids of the device.
            WriteLine(descF + "SetupDiBuildDriverInfoList ...")
            bResult = SetupDiBuildDriverInfoList(hDeviceInfoSet, m_DeviceInfoData, SPDIT_COMPATDRIVER)
            If bResult = False Then
                WriteLine("SetupDiBuildDriverInfoList: " & APIErrorByNumber())
                GoTo ForceInstall
            End If

            ' Pick the best driver in the list of drivers that was built.
            WriteLine(descF + "SetupDiCallClassInstaller ...")
            bResult = SetupDiCallClassInstaller(DIF_SELECTBESTCOMPATDRV, hDeviceInfoSet, m_DeviceInfoData)
            If bResult = False Then
                WriteLine("SetupDiBuildDriverInfoList: " & APIErrorByNumber())
                GoTo ForceInstall
            End If

            ' Get the selected driver node.
            ' Note: If this list does not contain any drivers, this call
            ' will fail with ERROR_NO_DRIVER_SELECTED.
            m_DriverInfoData.cbSize = Marshal.SizeOf(m_DriverInfoData)
            WriteLine(descF + "SetupDiGetSelectedDriver ...")
            bResult = SetupDiGetSelectedDriver(hDeviceInfoSet, m_DeviceInfoData, m_DriverInfoData)
            If bResult = False Then
                WriteLine("SetupDiGetSelectedDriver: " & APIErrorByNumber())
                GoTo ForceInstall
            End If

            'Dim dwRebootRequired As Long
            ' WriteLine(descF + "InstallSelectedDriver ...")
            'bResult = InstallSelectedDriver(IntPtr.Zero, hDeviceInfoSet, 0&, 0&, dwRebootRequired)
            'If bResult = False Then
            '    WriteLine("InstallSelectedDriver: " & APIErrorByNumber())
            '    GoTo ForceInstall
            'Else
            '    If (dwRebootRequired = DI_NEEDREBOOT) Or (dwRebootRequired = DI_NEEDRESTART) Then RebootRequired = True
            '    InstallModem = True
            '    WriteLine("InstallSelectedDriver: OK.")
            '    GoTo Cleanup
            'End If

            WriteLine(descF + "DiInstallDevice ...")
            bResult = DiInstallDevice(IntPtr.Zero, hDeviceInfoSet, m_DeviceInfoData, m_DriverInfoData, 0, RebootRequired)
            If bResult = False Then
                DisplayError("DiInstallDevice", "classGuid:" + m_DeviceInfoData.classGuid.ToString)
                GoTo ForceInstall
            Else
                WriteLine("DiInstallDevice: OK.")
                GoTo Cleanup
            End If

ForceInstall:

            'Usiamo UpdateDriverForPlugAndPlayDevices(anche con una chiara indicazione del file inf)
            'Evitare di utilizzare questo metodo, perché aggiorna il driver per tutti i dispositivi con lo stesso Hardware_ID, che è inutile installato
            'puramente di copertura, una volta fatto
            WriteLine(descF + "UpdateDriverForPlugAndPlayDevices ...")
            bResult = UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, Hardware_ID, INF_File, INSTALLFLAG_FORCE, RebootRequired)

            If bResult = False Then
                WriteLine(descF + "UpdateDriverForPlugAndPlayDevices " & APIErrorByNumber())
                bRemove = True
                GoTo Cleanup
            Else
                InstallModem = True
                WriteLine(descF + "UpdateDriverForPlugAndPlayDevices OK.")
                GoTo Cleanup
            End If

Cleanup:
            If bRemove Then
                ' Delete Device Instance that was registered using SetupDiRegisterDeviceInfo. May through an error if Device not registered -- who cares??
                WriteLine(descF + "SetupDiCallClassInstaller ...")
                SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, m_DeviceInfoData)
            End If
            WriteLine(descF + "SetupDiDestroyDeviceInfoList ...")
            SetupDiDestroyDeviceInfoList(hDeviceInfoSet.ToInt32)

            Return True

        Catch ex As Exception
            WriteLine(descF + "error: ..." + ex.Message)
            ClassMsgErr.LogErr("ClassMdmInstall", "InstallDriver", ex, True)
        End Try

    End Function













    Public Function InstallDriver(HardwareId As String, DriverInfFile As String, ByRef needReboot As Boolean) As Boolean
        Try
            Const descF As String = "InstallDriver: "
            Dim Hr As Boolean = True
            Dim ClassGUID As New Guid
            Dim className As New StringBuilder(256)
            Dim DeviceInfoSet As IntPtr
            Dim DeviceInfoData As SP_DEVINFO_DATA

            needReboot = False

            WriteLine(descF + "SetupDiGetINFClass: HardwareId=" + HardwareId + ",  FileInf= """ + DriverInfFile + """")

            If SetupDiGetINFClass(DriverInfFile, ClassGUID, className, 256, 0) = False Then
                DisplayError("SetupDiGetINFClass", "className:" + className.ToString)
                Return False
            End If

            WriteLine(descF + " ok, className=""" + className.ToString + """,  classGUID=""" + ClassGUID.ToString.ToUpper + """")

            'Create the container for the to-be-created Device Information Element.
            WriteLine(descF + "SetupDiCreateDeviceInfoList ...")
            DeviceInfoSet = SetupDiCreateDeviceInfoList(ClassGUID, IntPtr.Zero)
            If DeviceInfoSet = New IntPtr(-1) Then 'INVALID_HANDLE_VALUE 
                DisplayError("SetupDiCreateDeviceInfoList", "INVALID_HANDLE_VALUE")
                Return False
            End If

            ' Now create the element.  Use the Class GUID and Name from the INF file.
            DeviceInfoData.cbSize = Marshal.SizeOf(DeviceInfoData)

            WriteLine(descF + "DeviceInfoData.cbSize=" + DeviceInfoData.cbSize.ToString)
            If Not SetupDiCreateDeviceInfo(DeviceInfoSet, className.ToString, ClassGUID, "@mdmhayes.inf,%m2700%;Modem M5", IntPtr.Zero, DICD_GENERATE_ID, DeviceInfoData) Then
                DisplayError("SetupDiCreateDeviceInfo", "classGuid:" + DeviceInfoData.classGuid.ToString)
                Return False
            End If

            WriteLine(descF + "DeviceInfoData.classGuid=""" + DeviceInfoData.classGuid.ToString + """")
            WriteLine(descF + "DeviceInfoData.propertyId=""" + DeviceInfoData.propertyId.ToString + """")
            WriteLine(descF + "DeviceInfoData.cbSize=""" + DeviceInfoData.cbSize.ToString + """")

            WriteLine(descF + "SetupDiSetDeviceRegistryProperty: " + HardwareId)
            If Not SetupDiSetDeviceRegistryProperty(DeviceInfoSet, DeviceInfoData, SPDRP_HARDWAREID, HardwareId, (HardwareId.Length + 2) * ClassMdmInstall.GetcharSize()) Then
                DisplayError("SetupDiSetDeviceRegistryProperty", "INVALID_HANDLE_VALUE")
                SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)
                Return False
            End If

            Dim InstallParams As New SP_DEVINSTALL_PARAMS
            InstallParams.cbSize = Marshal.SizeOf(InstallParams)
            InstallParams.FlagsEx = DI_FLAGSEX_ALLOWEXCLUDEDDRVS Or DI_FLAGSEX_ALWAYSWRITEIDS
            InstallParams.Flags = DI_QUIETINSTALL Or DI_ENUMSINGLEINF
            '   InstallParams.DriverPath = New StringBuilder
            ' InstallParams.DriverPath.Append(DriverInfFile)

            'WriteLine(descF + "SetupDiGetDeviceInstallParams")
            'If SetupDiGetDeviceInstallParams(DeviceInfoSet, DeviceInfoData, InstallParams) = False Then
            '    DisplayError("SetupDiGetDeviceInstallParams", "classGuid: " + DeviceInfoData.classGuid.ToString.ToUpper)
            '    SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)
            '    Return False
            'End If


            WriteLine(descF + "SetupDiSetDeviceInstallParams")
            If SetupDiSetDeviceInstallParams(DeviceInfoSet, DeviceInfoData, InstallParams) = False Then
                DisplayError("SetupDiSetDeviceInstallParams", "classGuid: " + DeviceInfoData.classGuid.ToString.ToUpper)
                SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)
                Return False
            End If

            Dim DriverInfoData As New SP_DRVINFO_DATA

            WriteLine(descF + "SetupDiBuildDriverInfoList")
            If SetupDiBuildDriverInfoList(DeviceInfoSet, DeviceInfoData, SPDIT_COMPATDRIVER) = False Then
                DisplayError("SetupDiBuildDriverInfoList", "classGuid: " + DeviceInfoData.classGuid.ToString.ToUpper)
                SetupDiDestroyDriverInfoList(DeviceInfoSet, DeviceInfoData, SPDIT_COMPATDRIVER)
                Hr = False
            End If

            ' Use first best driver (since specified by inf file)
            WriteLine(descF + "SetupDiEnumDriverInfo")
            If SetupDiEnumDriverInfo(DeviceInfoSet, DeviceInfoData, SPDIT_COMPATDRIVER, 0, DriverInfoData) Then
                SetupDiSetSelectedDriver(DeviceInfoSet, DeviceInfoData, DriverInfoData)
            End If

            WriteLine(descF + "SetupDiCallClassInstaller")
            If SetupDiCallClassInstaller(DIF_REGISTERDEVICE, DeviceInfoSet, DeviceInfoData) = False Then
                DisplayError("SetupDiCallClassInstaller", "classGuid: " + DeviceInfoData.classGuid.ToString.ToUpper)
                Hr = False
            End If

            ' TODO Allow non interactive mode for drivers already contained in %SystemRoot%\inf directory

            ' BOOL PreviousMode = SetupSetNonInteractiveMode(TRUE);

            If (Hr = True) Then

                WriteLine(descF + "DiInstallDevice")
                If (DiInstallDevice(IntPtr.Zero, DeviceInfoSet, DeviceInfoData, DriverInfoData, 0, needReboot) = False) Then
                    DisplayError("SetupDiCallClassInstaller", "classGuid:" + DeviceInfoData.classGuid.ToString)
                    ' Ensure that the device entry in \ROOT\ENUM\ will be removed...
                    SetupDiRemoveDevice(DeviceInfoSet, DeviceInfoData)
                    Hr = False
                End If
            End If

            'SetupSetNonInteractiveMode(PreviousMode)

            SetupDiDestroyDeviceInfoList(DeviceInfoSet.ToInt32)

            Return Hr

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassMdmInstall", "InstallDriver", ex, True)
        End Try

    End Function

End Class
