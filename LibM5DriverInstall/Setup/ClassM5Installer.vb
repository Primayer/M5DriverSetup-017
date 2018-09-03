Imports LibPaths
Imports LibDownloadMcp


Public Class ClassM5Installer
#If MCPVER = "PRO" Then
    Public Const WRITE_USB_FLASGS As Boolean = True
#Else
    Public Const WRITE_USB_FLASGS As Boolean = False
#End If

    Private _txtLog As TextBox
    Public WithEvents mdmInst As ClassMdmInstall
    Public RasBook As New LibRasBook.ClassRasBook

    Private pic As GroupBox
    Private btnInst As Button
    Private btnCancel As Button
    Private _exit As Boolean = False
    Private _isRunning As Boolean = False
    Private btnUminstall As Button

    Public Event InstallFinish()
    Public Event UninstallFinish()
    Public Event TreeviewPortUpdated(ByVal nPortInstalled As Integer)

    Public ReadOnly Property IsRunning() As Boolean
        Get
            Return Me._isRunning
        End Get
    End Property

    Public Sub LogAppend(ByVal s As String)
        Try
            Safe.Textbox.AppendText(s, Me._txtLog)
            Me.mdmInst.PrintLog(s)
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "LogAppend", ex, True)
        End Try
    End Sub

    Public Sub LogReset()
        Try
            Safe.Control.Text("", Me._txtLog)
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "LogReset", ex, True)
        End Try
    End Sub


    Private Function GetPathFileInf(UsbRaw As String, PID As String, ByRef inf_ls As String) As Boolean
        Try

            ' crea nome file inf del modem da installare a seconda del PID:
            Select Case PID
                Case "7018"
                    ' seleziono il file inf con la velocità adeguata: LS
                    inf_ls = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileMdmInf_LS)

                Case "7019"
                    ' seleziono il file inf con la velocità adeguata: HS
                    inf_ls = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileMdmInf_HS)

                Case "6015"
                    ' old version
                    inf_ls = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileMdmInf_LS)

                Case Else
                    Return False
            End Select

            Return True

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "GetPathFileInf", ex, True)
        End Try
    End Function

    Private Function GetPathFileInf_old(UsbRaw As String, ByRef inf_ls As String) As Boolean
        Try
            'Dim w As String = Environment.GetEnvironmentVariable("windir")
            'inf = My.Computer.FileSystem.CombinePath(w, "inf")
            'inf = My.Computer.FileSystem.CombinePath(inf, "mdmhayes.inf")
            'inf = "C:\temp\Driver\m5\M5-MDM-HEMINA.inf"
            'inf = "C:\Users\msimonato.HEMINA\Desktop\Installer\Installer\Drivers\MODEM\mdmhayes.inf"
            'inf = "C:\temp\Driver\m5\M5-MDM-LS.inf"
            'inf = "C:\temp\Driver\m5\M5-MDM-LS.inf"


            '                    Product IsRX21A IsRX64M IsHighSpeed
            Dim lp As New List(Of Board_desc) From {
                New Board_desc("HEMINA-MV110-HW00", 1, 0, 0),
                New Board_desc("HEMINA-MV252-HW01", 1, 0, 0),
                New Board_desc("HEMINA-MV210-HW02", 1, 1, 0),
                New Board_desc("HEMINA-MV800-HW03", 1, 0, 0),
                New Board_desc("HEMINA-MV801-HW04", 1, 0, 0),
                New Board_desc("HEMINA-MV255-HW05", 1, 1, 1),
                New Board_desc("HEMINA-MV810-HW06", 1, 0, 0),
                New Board_desc("HEMINA-MV311-HW07", 0, 1, 0)
                           }

            Dim idx As Integer = -1
            For i As Integer = 0 To lp.Count - 1
                If lp(i).Product = UsbRaw Then
                    idx = i
                    Exit For
                End If
            Next

            If idx >= 0 Then

                If lp(idx).IsHighSpeed = 1 Then
                    ' a seconda della descrizione raw usb, seleziono il file inf con la velocità adeguata: HS
                    inf_ls = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileMdmInf_HS)
                Else
                    ' a seconda della descrizione raw usb, seleziono il file inf con la velocità adeguata: LS
                    inf_ls = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileMdmInf_LS)
                End If

            Else
                Return False
            End If

            Return True

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "GetPathFileInf", ex, True)
        End Try
    End Function

    Class Board_desc
        Public Product As String = ""
        Public IsRX21A As Byte = 0
        Public IsRX64M As Byte = 0
        Public IsHighSpeed As Byte = 0

        Public Sub New(Product As String, IsRX21A As Byte, IsRX64M As Byte, IsHighSpeed As Byte)
            Me.Product = Product
            Me.IsRX21A = IsRX21A
            Me.IsRX64M = IsRX64M
            Me.IsHighSpeed = IsHighSpeed

        End Sub
    End Class


    Public Function PathFtdi() As String
        Try

            Return My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileFTDIM5.Replace(".zip", ""))
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "PathFtdiHS", ex, False)
            Return ""
        End Try
    End Function


    Public Function PathFtdiHS_old() As String
        Try

            ' Return My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileFTDIM5_HS)
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "PathFtdiHS", ex, False)
            Return ""
        End Try
    End Function


    Public Function PathFtdiLS_old() As String
        Try

            '   Return My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5Driver, ClassEstraiRisorsa.FileFTDIM5_LS)




            Dim p As String = ""

            p = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.CreaDirAppdataDownload, "FTDI")
            Dim l As New List(Of String)


            For Each s As String In My.Computer.FileSystem.GetFiles(p)
                Dim f As String = My.Computer.FileSystem.GetName(System.IO.Path.GetFileName(s))
                If f.ToUpper.StartsWith("CDM_V") AndAlso System.IO.Path.GetExtension(s).ToLower = ".exe" AndAlso (f.IndexOf("."c) >= 0) Then
                    l.Add(s)
                End If
            Next

            'CDM_v2.12.28.exe
            Dim verList As New List(Of Long)

            If l.Count > 0 Then
                For i As Integer = 0 To l.Count - 1
                    Dim ver0 As String = ""
                    Dim ver1 As String = ""
                    Dim ver2 As String = ""
                    Dim name As String = ""
                    name = My.Computer.FileSystem.GetName(l(i))
                    Dim s() As String
                    s = name.Split("."c)
                    If s IsNot Nothing AndAlso s.Length >= 4 Then
                        ver0 = s(0).Substring("CDM_V".Length)
                        ver1 = s(1)
                        ver2 = s(2)
                        Dim v0 As Integer = 0
                        Dim v1 As Integer = 0
                        Dim v2 As Integer = 0
                        If Integer.TryParse(ver0, v0) AndAlso Integer.TryParse(ver1, v1) AndAlso Integer.TryParse(ver2, v2) Then
                            Dim numStr As String = ver0 + ver1 + ver2
                            Dim num As Long = 0
                            If Long.TryParse(numStr, num) Then
                                verList.Add(num)
                            End If
                        End If

                    End If
                Next
                If verList.Count > 0 AndAlso verList.Count = l.Count Then
                    ' trova il masssimo
                    Dim max As Long = verList(0)
                    Dim imax As Integer = 0
                    For i As Integer = 1 To verList.Count - 1
                        If verList(i) > max Then
                            max = verList(i)
                            imax = i
                        End If
                    Next
                    Return l(imax)
                End If
                Return ""
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "PathFtdiLS", ex, False)
            Return ""
        End Try
    End Function


    Public Function PathFtdi____() As String
        Try
            Dim foundOnLocalDir As Boolean = False

            Dim p As String = ""



            p = My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.CreaDirAppdataDownload, "FTDI")

            For Each s As String In My.Computer.FileSystem.GetFiles(p)
                Dim f As String = My.Computer.FileSystem.GetName(System.IO.Path.GetFileName(s))
                If f.ToUpper.StartsWith("CDM_V") AndAlso System.IO.Path.GetExtension(s).ToLower = ".exe" Then
                    foundOnLocalDir = True
                    Return s
                End If
            Next

            'Dim c As String = ""

            'c = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

            'If foundOnLocalDir = False Then

            '    p = My.Computer.FileSystem.CombinePath(c, "\Mcp\Download\")
            '    For Each s As String In My.Computer.FileSystem.GetFiles(p)
            '        Dim f As String = My.Computer.FileSystem.GetName(System.IO.Path.GetFileName(s))
            '        If f.ToUpper.StartsWith("CDM_V") AndAlso System.IO.Path.GetExtension(s).ToLower = ".exe" Then
            '            foundOnLocalDir = True
            '            Return s
            '        End If
            '    Next 
            '    Return ""
            'End If

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "PathFtdi", ex, False)
            Return ""
        End Try
    End Function

    Public Function StartInstall(ByVal innstallFTDI As Boolean, ByVal btnInst As Button, ByVal btnCancel As Button, ByVal btnUminstall As Button) As Boolean
        Try
            Me.btnInst = btnInst
            Me.btnCancel = btnCancel
            Me.btnUminstall = btnUminstall
            Me._exit = False

            Safe.Control.Enabled(False, Me.btnInst)
            Safe.Control.Enabled(False, Me.btnUminstall)

            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf threadInstall, innstallFTDI)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "StartInstall", ex, True)
        End Try

    End Function

    Enum TipoUnistall
        Modem
        Port
    End Enum

    Class classUnInMdm
        Public Port As String = ""
        Public tipo As TipoUnistall
        Public btnInst As Button
        Public btnUminstall As Button
    End Class

    Public Sub StartUnistallModem(ByVal Port As String, ByVal tipo As TipoUnistall, ByVal btnInst As Button, ByVal btnUminstall As Button)
        Try
            Dim p As New classUnInMdm

            p.Port = Port
            p.tipo = tipo
            p.btnInst = btnInst
            p.btnUminstall = btnUminstall
            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _UnistallModem, p)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub _UnistallModem(ByVal o As Object)
        Try
            Dim p As classUnInMdm = CType(o, classUnInMdm)

            UnistallModem(p.Port, p.tipo, p.btnInst, p.btnUminstall)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub UnistallModem(ByVal Port As String, ByVal tipo As TipoUnistall, ByVal btnInst As Button, ByVal btnUminstall As Button)
        Try
            Me.btnInst = btnInst
            Me.btnUminstall = btnUminstall

            Safe.Control.Enabled(False, Me.btnInst)
            Safe.Control.Enabled(False, Me.btnUminstall)

            Dim un As New ClassUnistall

            LogReset()

            If tipo = TipoUnistall.Modem Then
                LogAppend("Unistalling M5 device on port " + Port + " ..." + vbCrLf)
                If Me.mdmInst.ModemExist(Port) Then
                    If Me.mdmInst.UninstallModem(Port) Then
                        LogAppend("M5 device successfully uninstalled!" + vbCrLf)
                    Else
                        LogAppend("Error unistalling M5 device!" + vbCrLf)
                    End If
                Else
                    LogAppend("Not found modem on port " + Port + "! " + vbCrLf)
                End If
            Else
                LogAppend("Unistalling M5 port " + Port + " ..." + vbCrLf)
                If Me.mdmInst.UninstallPort(Port) Then
                    LogAppend("M5 port successfully uninstalled!" + vbCrLf)
                Else
                    LogAppend("Error unistalling M5 port!" + vbCrLf)
                End If
            End If


            RaiseEvent UninstallFinish()

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "StartUnistall", ex, True)
        Finally
            Safe.Control.Enabled(True, Me.btnInst)
            Safe.Control.Enabled(True, Me.btnUminstall)
        End Try
    End Sub

    Public Function StopInstall() As Boolean
        Try
            Me._exit = True
            LogAppend(vbCrLf + "Cancel ...")
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "StopInstall", ex, True)
        End Try
    End Function

    Private Sub wait(ByVal ms As Integer)
        Try
            Dim _sw As New Stopwatch
            _sw.Reset()
            _sw.Start()
            Do
                If Me._exit = True Then Exit Do
                If _sw.Elapsed.TotalMilliseconds >= ms Then
                    Exit Do
                End If

                System.Threading.Thread.Sleep(100)

            Loop
        Catch ex As Exception

        End Try
    End Sub

    Private Sub threadInstall(ByVal o As Object)
        Try
            Me._isRunning = True

            ' visualizza pic installazione
            Safe.Control.Visible(True, Me.pic)

            '  TopMostMessageBox.ShowNoBlock("thread start install")

            Dim innstallFTDI As Boolean = CType(o, Boolean)
            Dim ListFtdiInstalled As New List(Of ClassFtdiRegKey.ClassFTDIInstalled)
            Dim PortToIntsall As String = ""

            Dim portPlugged As String = ""

            'resetta il log
            LogReset()

            If Me._exit = True Then Exit Try

            If innstallFTDI Then
                LogAppend("Installing M5 driver ...")
                If InstallFTDI() = False Then
                    LogAppend(" : Error!")
                End If
                LogAppend(vbCrLf)
            End If

            Safe.Control.Enabled(True, Me.btnCancel)

            ' usb flag=1   -> quando si inserisce un nuovo sn non incrementa la porta
            If ClassM5Installer.WRITE_USB_FLASGS = True Then ClassFtdiRegKey.WriteUsbFlags(True)

            If Me._exit = True Then Exit Try

            ' wait FTDI com port installed
            If WaitAlmeno1FTDIInstalled(ListFtdiInstalled) = False Then
                Exit Try
            End If

            wait(2000)

            If Me._exit = True Then Exit Try

            Dim portsPluggedList As New List(Of String)

            ' wait plugged usb ftdi
            If WaitFTDIPluggedList(ListFtdiInstalled, portsPluggedList) Then
                System.Threading.Thread.Sleep(1000)
            End If

            If Me._exit = True Then Exit Try

            Dim listEsiti As New List(Of classEsitoInstall)

            For i As Integer = 0 To portsPluggedList.Count - 1

                portPlugged = portsPluggedList(i)

                If ClassFtdiRegKey.GetModemInstalled(portPlugged, True) Then
                    LogAppend(vbCrLf + "M5 device driver already installed on " + portPlugged + "!" + vbCrLf)
                    Continue For
                End If

                If Me._exit = True Then Exit Try

                ''''''''''''''' tolta il 03/05/2018 perchè inclusa nel driver M5''''''''''
                ' imposta emulazione modem per controllo di flusso hardware              '
                ' ClassFtdiRegKey.SetEmulationMode(portPlugged, True)                    '
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Dim Esito As New classEsitoInstall

                Esito.desvice = portPlugged

                Dim inf As String = ""

                Dim UsbRawObj As New ClassSerachHeminaUSBraw.classUSBraw
                Dim dtms As String = ""
                LogAppend("Read USB raw for " + portPlugged + " ... " + vbCrLf)
                Dim USB As New ClassSerachHeminaUSBraw.classUSB
                ' cerca la stringa USB raw che corrisponde alla porta plugged
                If GetUsbRaw(portPlugged, UsbRawObj, USB, dtms) Then
                    LogAppend(" USB raw for " + portPlugged + " is " + UsbRawObj.USBraw + " " + UsbRawObj.USBdescriprion + vbCrLf)

                    If GetPathFileInf(UsbRawObj.USBraw, UsbRawObj.PID, inf) = False Then
                        LogAppend("Error file inf!" + vbCrLf)
                        Continue For
                    End If
                Else
                    ' non trovato tra i dispositivi HEMINA

                    LogAppend(portPlugged + " is not a M5 device!" + vbCrLf)
                    Continue For
                End If

                If Me._exit = True Then Exit Try


                ' Install file inf modem
                LogAppend(vbCrLf + "Installing device driver on " + portPlugged + " ..." + vbCrLf)
                LogAppend(inf + vbCrLf + vbCrLf)

                If Me._exit = True Then Exit Try

                Dim speed As ClassMdmInstall.EnumM5Speed = ClassMdmInstall.EnumM5Speed.LowSpeed
                Select Case UsbRawObj.M5Speed
                    Case ClassMdmInstall.EnumM5Speed.LowSpeed.ToString
                        speed = ClassMdmInstall.EnumM5Speed.LowSpeed
                    Case ClassMdmInstall.EnumM5Speed.HighSpeed.ToString
                        speed = ClassMdmInstall.EnumM5Speed.HighSpeed
                    Case Else
                        speed = ClassMdmInstall.EnumM5Speed.LowSpeed
                End Select

                If Me.mdmInst.InstallDeviceAndDriver(inf, ClassMdmInstall.GetHardawrareID(speed), portPlugged) Then
                    LogAppend(vbCrLf + "M5 driver installed successfull on " + portPlugged + vbCrLf + vbCrLf)
                    ' TopMostMessageBox.ShowNoBlock("M5 driver installed successfull!" + vbCrLf + vbCrLf + "Disconnect and reconnect the M5 USB connector!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Esito.Esito = True
                    Esito.Errdesc = "M5 driver installed successfull on " + portPlugged
                Else
                    Esito.Esito = False
                    LogAppend(vbCrLf + "Error installing device driver on port " + portPlugged + "!" + vbCrLf + vbCrLf)
                    Esito.Errdesc = "Error installing device driver on port " + portPlugged + "!"
                    TopMostMessageBox.ShowNoBlock(Esito.Errdesc, "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

                listEsiti.Add(Esito)

            Next

            Dim sOK As String = ""
            Dim sFail As String = ""
            For i As Integer = 0 To listEsiti.Count - 1
                If listEsiti(i).Esito = True Then
                    sOK = sOK + listEsiti(i).Errdesc + vbCrLf
                Else
                    sFail = sFail + listEsiti(i).Errdesc + vbCrLf
                End If
            Next

            If sOK <> "" Then
                TopMostMessageBox.ShowNoBlock(sOK + vbCrLf + vbCrLf + "Disconnect and reconnect the M5 USB connector!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "threadInstall", ex, True)
        Finally
            ' nascondi pic installazione
            Safe.Control.Visible(False, Me.pic)
            Safe.Control.Enabled(True, Me.btnInst)
            Safe.Control.Enabled(True, Me.btnUminstall)
            Safe.Control.Enabled(False, Me.btnCancel)
            If Me._exit = True Then LogAppend("Cancelled!")
            RaiseEvent InstallFinish()
            Me._isRunning = False
        End Try
    End Sub

    Class classEsitoInstall
        Public Esito As Boolean = False
        Public Errdesc As String = ""
        Public desvice As String = ""
    End Class

    Private Function GetUsbRawFtdi(portPlugged As String, ByRef UsbRaw As String, ByRef dtms As String) As Boolean
        Try
            UsbRaw = ""

            Dim ftdi As New LibM5Ftdi.ClassFTDI
            Dim l As List(Of LibM5Ftdi.ClassFTDI.ClassUsbM5)



            Dim _sw As New Stopwatch
            _sw.Reset()
            _sw.Start()

            Do
                l = ftdi.SearchM5(dtms)
                If l.Count > 0 Then
                    Exit Do
                End If

                If _sw.Elapsed.TotalMilliseconds >= 10000 Then
                    Exit Do
                End If
                System.Threading.Thread.Sleep(400)
            Loop


            For i As Integer = 0 To l.Count - 1
                If portPlugged = l(i).ComPort Then
                    UsbRaw = l(i).Description
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "GetUsbRaw", ex, True)
        End Try
    End Function

    Private Function GetUsbRaw(portPlugged As String, ByRef UsbRaw As ClassSerachHeminaUSBraw.classUSBraw, ByRef USB As ClassSerachHeminaUSBraw.classUSB, ByRef dtms As String) As Boolean
        Try
            UsbRaw = New ClassSerachHeminaUSBraw.classUSBraw

            Dim SerachHeminaUSBraw As New ClassSerachHeminaUSBraw
            USB = New ClassSerachHeminaUSBraw.classUSB

            If SerachHeminaUSBraw.SearchHeminaUSB(SerachHeminaUSBraw.FullPath_M5UsbListexe, SerachHeminaUSBraw.NomeFileToSave, USB) Then

                For i As Integer = 0 To USB.DeviceList.Count - 1
                    If portPlugged.Trim = USB.DeviceList(i).COM.Trim Then
                        UsbRaw = USB.DeviceList(i)
                        Return True
                    End If
                Next

            End If

            Return False

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "GetUsbRaw", ex, True)
        End Try
    End Function

    Private Function WaitFTDIInstalled(ByRef PortName As String) As Boolean
        Try
            Dim l As New List(Of ClassFtdiRegKey.ClassFTDIInstalled)
            Dim errore As Boolean = False

            LogAppend(vbCrLf + "Search for installed M5 USB device ..." + vbCrLf)

            Do
                If Me._exit = True Then Exit Try

                System.Threading.Thread.Sleep(300)

                l = ClassFtdiRegKey.SearchFtdiInstalled(errore)

                If errore Then
                    TopMostMessageBox.ShowNoBlock("Error FTDI port not installed!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Do
                End If

                If l.Count > 0 Then
                    PortName = l(0).PortName
                    For i As Integer = 0 To l.Count - 1
                        LogAppend("  Found port installed: " + (i + 1).ToString + "/" + l.Count.ToString + " : " + l(i).PortName + vbCrLf)
                    Next
                    Return True
                End If


            Loop
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "WaitFTDIInstalled", ex, True)
        End Try

    End Function

    Private Function WaitAlmeno1FTDIInstalled(ByRef ListFtdiInstalled As List(Of ClassFtdiRegKey.ClassFTDIInstalled)) As Boolean
        Try
            Dim _ListFtdiInstalled As New List(Of ClassFtdiRegKey.ClassFTDIInstalled)
            ' Dim _ListFtdi6015_NoM5SpeedInstalled As New List(Of ClassFtdiRegKey.ClassFTDIInstalled)
            Dim errore As Boolean = False

            LogAppend(vbCrLf + "Search for installed M5 USB device ..." + vbCrLf)

            Do

                If Me._exit = True Then Exit Try

                System.Threading.Thread.Sleep(300)

                _ListFtdiInstalled = ClassFtdiRegKey.SearchFtdiInstalled(errore)

                ' _ListFtdi6015_NoM5SpeedInstalled = ClassFtdiRegKey.SearchFtdi6015_NoM5SpeedInstalled(errore)

                If errore Then
                    '  TopMostMessageBox.ShowNoBlock("Error FTDI port not installed!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ' Exit Do
                End If

                If _ListFtdiInstalled.Count > 0 Then
                    ListFtdiInstalled = _ListFtdiInstalled

                    'If _ListFtdi6015_NoM5SpeedInstalled IsNot Nothing AndAlso _ListFtdi6015_NoM5SpeedInstalled.Count > 0 Then
                    '    ListFtdiInstalled.AddRange(_ListFtdi6015_NoM5SpeedInstalled.ToArray)
                    'End If

                    For i As Integer = 0 To _ListFtdiInstalled.Count - 1
                        LogAppend("  Found port installed: " + (i + 1).ToString + "/" + _ListFtdiInstalled.Count.ToString + " : " + _ListFtdiInstalled(i).PortName + " [VID_" + _ListFtdiInstalled(i).VID + " PID_" + _ListFtdiInstalled(i).PID + "]" + vbCrLf)
                    Next
                    Return True
                End If


            Loop
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "WaitFTDIInstalled", ex, True)
        End Try

    End Function

    Private Function WaitFTDIPlugged(ByVal PortName As String) As Boolean
        Try

            LogAppend(vbCrLf + "Waiting for plugging M5 USB ..." + vbCrLf)

            Do
                If Me._exit = True Then Exit Try

                If Me.mdmInst.PortExist(PortName, False) Then
                    LogAppend("  Found port plugged: " + PortName + vbCrLf)
                    LogAppend(vbCrLf)
                    Return True
                End If

                System.Threading.Thread.Sleep(100)
            Loop

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "WaitFTDIPlugged", ex, True)
        End Try

    End Function

    Private Function WaitFTDIPlugged(ByVal PortNames As List(Of String), ByRef portPlugged As String) As Boolean
        Try

            LogAppend(vbCrLf + "Waiting for plugging M5 USB ..." + vbCrLf)
            portPlugged = ""

            Do
                For i As Integer = 0 To PortNames.Count - 1

                    If Me._exit = True Then Exit Try

                    If Me.mdmInst.PortExist(PortNames(i), False) Then
                        portPlugged = PortNames(i)
                        LogAppend("  Found port plugged: " + portPlugged + vbCrLf)
                        Return True
                    End If
                    System.Threading.Thread.Sleep(16)
                Next

                System.Threading.Thread.Sleep(100)
            Loop

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "WaitFTDIPlugged", ex, True)
        End Try

    End Function

    Private Function WaitFTDIPluggedList(ByVal ListFtdiInstalled As List(Of ClassFtdiRegKey.ClassFTDIInstalled),
                                         ByRef portPluggedList As List(Of String)) As Boolean
        Try

            LogAppend(vbCrLf + "Waiting for plugging M5 USB ..." + vbCrLf)

            If portPluggedList Is Nothing Then
                portPluggedList = New List(Of String)
            Else
                portPluggedList.Clear()
            End If

            Do
                For i As Integer = 0 To ListFtdiInstalled.Count - 1

                    If Me._exit = True Then Exit Try

                    If Me.mdmInst.PortExist(ListFtdiInstalled(i).PortName, False) Then
                        portPluggedList.Add(ListFtdiInstalled(i).PortName)
                        LogAppend("  Found port plugged: " + ListFtdiInstalled(i).PortName + vbCrLf)
                    End If
                    System.Threading.Thread.Sleep(16)
                Next

                If portPluggedList.Count > 0 Then
                    LogAppend(vbCrLf)
                    Return True
                End If

                System.Threading.Thread.Sleep(100)
            Loop


        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "WaitFTDIPluggedList", ex, True)
        End Try

    End Function

    Private updateObj As ClassUpdate
    Private manualEvent As System.Threading.ManualResetEvent
    Private Sub downloadFinishLogSAved(ByVal nFilesToDownload As Integer, ByVal NfilesDownloaded As Integer)
        Try
            RemoveHandler Me.updateObj.DownLoadFilesListFinishLogSaved, AddressOf downloadFinishLogSAved
            Dim s As String = ""
            Dim F As String = My.Computer.FileSystem.CombinePath(ClassPATH.CreaDirAppdataDownload, ClassPATH.NomeFileLogDownload)
            s = My.Computer.FileSystem.ReadAllText(F)
            'Safe.Control.Text(s, Me.TextBox1)
            manualEvent.Set()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "downloadFinish", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Function InstallFTDI() As Boolean
        Try
            '  Dim ftdiPath_LS As String = PathFtdiLS().ToUpper.Replace(".ZIP", "") + "\dp-chooser.exe"
            'Dim ftdiPath_LS1 As String = PathFtdiLS().ToUpper.Replace(".ZIP", "") + "\M5-MDM-LS.inf"
            'Dim ftdiPath_LS2 As String = PathFtdiLS().ToUpper.Replace(".ZIP", "") + "\M5_LS_bus.inf"


            'Dim ftdiPath_LS As String = PathFtdiLS().ToUpper.Replace(".ZIP", "")
            'Dim ftdiPath_HS As String = PathFtdiHS().ToUpper.Replace(".ZIP", "")



            'Return ClassUnistallFtdi.InstallDriverpPnputil(ftdiPath_LS1)

            'Return _InstallFTDI("pnputil.exe", "-a -i " + ftdiPath_LS1) And _InstallFTDI("pnputil.exe", "-a -i " + ftdiPath_LS2)

            '  Return _InstallFTDI(ftdiPath_LS, "") And _InstallFTDI(ftdiPath_HS, "")

            ' Return LaunchProcess.LaunchFile(ftdiPath_LS, True) And LaunchProcess.LaunchFile(ftdiPath_HS, True)

            Return ProcessStart("pnputil.exe", "-i -a " + ClassEstraiRisorsa.FileM5_LS_bus_inf, PathFtdi, _txtLog, False) AndAlso
                   ProcessStart("pnputil.exe", "-i -a " + ClassEstraiRisorsa.FileM5_LS_port_inf, PathFtdi, _txtLog, False) AndAlso
                   ProcessStart("pnputil.exe", "-i -a " + ClassEstraiRisorsa.FileM5_HS_bus_inf, PathFtdi, _txtLog, False) AndAlso
                   ProcessStart("pnputil.exe", "-i -a " + ClassEstraiRisorsa.FileM5_HS_port_inf, PathFtdi, _txtLog, False)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "InstallFTDI", ex, True)
        End Try

    End Function

    Public Function ProcessStart(nomeFileProcesso As String, Arguments As String, WorkingDirectory As String, txtLogUnistall As TextBox, clearLog As Boolean) As Boolean
        Try
            If clearLog Then Safe.Control.Text("", txtLogUnistall)

            ' prepara processo  'https://www.dotnetperls.com/redirectstandardoutput
            Dim pi As New ProcessStartInfo()
            pi.CreateNoWindow = True
            pi.WindowStyle = ProcessWindowStyle.Hidden
            pi.UseShellExecute = False
            pi.RedirectStandardOutput = True
            pi.FileName = nomeFileProcesso
            pi.Arguments = Arguments
            If WorkingDirectory <> "" Then
                pi.WorkingDirectory = WorkingDirectory
            End If

            ' lancia processo

            Dim p As Process = Process.Start(pi)

            p.StartInfo.Verb = "runas"

            ' p = Process.Start(pi)

            ' legge stream di output
            Using rd As New System.IO.StreamReader(p.StandardOutput.BaseStream)
                Dim s As String = rd.ReadToEnd
                If txtLogUnistall IsNot Nothing Then
                    Safe.Textbox.AppendText(s, txtLogUnistall)
                End If
            End Using

            ' attende che finisca
            p.WaitForExit()

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "ProcessStart", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Private Function _InstallFTDI(ftdiPath As String, Arguments As String) As Boolean
        Try
            Dim startInfo As New ProcessStartInfo(ftdiPath)

            startInfo.WindowStyle = ProcessWindowStyle.Normal

            startInfo.UseShellExecute = True
            startInfo.Arguments = Arguments

            Dim p As New Process

            p.StartInfo.Verb = "runas"

            If Me._exit = True Then Exit Try

            p = Process.Start(startInfo)

            If p IsNot Nothing Then
                p.WaitForExit()

                Dim result As Int32 = p.ExitCode
                If result <> 0 Then
                    LogAppend(ftdiPath + vbCrLf + "ExitCode: " + p.ExitCode.ToString)
                End If

            Else
                MessageBox.Show("FTDI Setup driver error!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Return True
        Catch ex As Exception

        End Try
    End Function



    Public Function InstallFTDI_old() As Boolean
        Try
            'Dim ftdiPath As String = PathFtdiLS()

            'If ftdiPath = "" Then
            '    If MessageBox.Show("Not found FTDI installer!" + vbCrLf + vbCrLf + "Would you like to download it?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            '        ' download
            '        Me.updateObj = New ClassUpdate(LibPaths.ClassPATH.CreaDirAppdataDownload)

            '        AddHandler Me.updateObj.DownLoadFilesListFinishLogSaved, AddressOf downloadFinishLogSAved
            '        Me.updateObj.StartDownload(Nothing)
            '        Me.manualEvent = New System.Threading.ManualResetEvent(False)

            '        manualEvent.WaitOne()

            '    Else
            '        Return False
            '    End If

            'End If

            'ftdiPath = PathFtdiLS()
            'If ftdiPath = "" Then
            '    MessageBox.Show("FTDI path error!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Return False
            'End If

            'Dim startInfo As New ProcessStartInfo(ftdiPath)

            'startInfo.WindowStyle = ProcessWindowStyle.Normal

            'startInfo.UseShellExecute = True

            'Dim p As Process = Process.Start(startInfo)

            'If Me._exit = True Then Exit Try

            'If p IsNot Nothing Then
            '    p.WaitForExit()
            'Else
            '    MessageBox.Show("FTDI Setup driver error!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End If

            Return True

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassM5Installer", "InstallFTDI", ex, True)
        End Try

    End Function

    Public Sub New(ByVal txtLog As TextBox, ByVal pic As GroupBox, lblReadingRawString As Label, txtUsbRaw As TextBox, btnRefresh As Button, tablLytPanelUnistall As TableLayoutPanel)
        Try
            Me._txtLog = txtLog
            Me.pic = pic
            mdmInst = New ClassMdmInstall(Nothing, ClassPATH.CreaDirAppdataDownload, lblReadingRawString, txtUsbRaw, btnRefresh, tablLytPanelUnistall)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub mdmInst_TreeviewPortUpdated(ByVal nPortInstalled As Integer) Handles mdmInst.TreeviewPortUpdated
        RaiseEvent TreeviewPortUpdated(nPortInstalled)
    End Sub

End Class
