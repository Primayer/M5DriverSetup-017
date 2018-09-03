Imports System.Runtime.InteropServices
Imports LibProxy

Public Class M5_Setup

    Public Const versione As String = "1.0.0.8"

    Private WithEvents m5Ints As ClassM5Installer
    Private _enableAutoInstall As Boolean = False

    Private Sub M5_Setup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'UsbNotification.RegisterUsbDeviceNotification(Me.Handle)

            Me.lblDescExtractor.Text = ""
            Me.txtUsbRaw.Text = ""
            Me.txtUsbRaw.Dock = DockStyle.Fill

            ClassDriverEvent.Register()

            'Me.usbEvnt = New USBEvent
            'Me.usbEvnt.Start(

            Me.TableLayoutPanel5.Dock = DockStyle.Fill

            If ClassUser.IsAdministrator Then
                Me.lbladmin.Text = ClassUser.GetName + " (Administrator)"

            Else
                Me.lbladmin.Text = ClassUser.GetName + " (NOT Administrator)"
            End If
            ' versione assembly
            'Dim ass As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom("LibM5DriverInstall.dll")
            ' Dim ver As Version = ass.GetName().Version
            Me.lblVer.Text = "Ver.: " + versione
            ' Me.lblVer.Text = "Ver.: " + System.Reflection.Assembly.GetAssembly(GetType(M5_Setup)).GetName.Version.ToString() ' + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString

            If LibM5DriverInstall.ClassUser.IsAdministrator Then
                Me.Text &= " - " + LibM5DriverInstall.ClassUser.GetName + " (Administrator)" + " " + My.Application.Info.Version.ToString
            Else
                Me.Text &= " " + LibM5DriverInstall.ClassUser.GetName + " " + My.Application.Info.Version.ToString
            End If

            Me.TableLayoutPanel4.Dock = DockStyle.Fill
            Me.TableLayoutPanel1.Dock = DockStyle.Fill

            Me.TableLayoutPanel1.Location = New Point(0, 0)
            Me.TableLayoutPanel1.Dock = DockStyle.Fill

            Me.lblWin.Text = My.Computer.Info.OSFullName + " " + ClassMdmInstall.Get32or64bit + vbCrLf + My.Computer.Info.OSVersion + vbCrLf + My.Computer.Info.InstalledUICulture.EnglishName

            ' usb flag=1   -> quando si inserisce un nuovo sn non incrementa la porta
            If ClassUser.IsAdministrator AndAlso ClassM5Installer.WRITE_USB_FLASGS = True Then ClassFtdiRegKey.WriteUsbFlags(True)

            ClassShowHideDevice.DevMgr_Show_NonPresent_Devices(True)

            Me.m5Ints = New ClassM5Installer(Me.TextBox1, Me.grpPicSetup, Me.lblReadingRawString, Me.txtUsbRaw, Me.btnRefresh, Me.TableLayoutPanel3)

            Dim drverFtdiOk As Boolean = False

            'FTDI driver Installed?
            If ClassFtdiRegKey.FtdiDriverInstalled = True Then
                '  Me.m5Ints.mdmInst.StartGetListModemInstalledAndFTDIPort(Me.TreeView1)
                drverFtdiOk = True
            Else
                ' Me.ParentForm.WindowState = FormWindowState.Minimized
                If MessageBox.Show("M5 driver Not Installed!" + vbCrLf + vbCrLf + "Would you like to install the driver?", "Attention!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    Me.m5Ints.LogAppend("Installing M5 driver ..." + vbCrLf)
                    If Me.m5Ints.InstallFTDI() = False Then
                        MessageBox.Show("M5 driver not installed correctly!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        drverFtdiOk = False
                        Me.m5Ints.LogAppend("M5 driver NOT installed correctly!")

                    Else
                        Me.m5Ints.LogAppend("M5 driver installed." + vbCrLf)
                        drverFtdiOk = True
                    End If
                Else
                    drverFtdiOk = False
                End If

                ' Me.ParentForm.WindowState = FormWindowState.Normal
            End If

            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf EstraiFIlesELeggiDevicesInstallati, drverFtdiOk)

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "FormClosing", ex, True)
        End Try
    End Sub

    Private Sub EstraiFIlesELeggiDevicesInstallati(o As Object)
        Try
            Dim drverFtdiOk As Boolean = CType(o, Boolean)


            ClassEstraiRisorsa.EstraiTutto(Me.lblDescExtractor)

            If drverFtdiOk Then
                Me.chkInstallFTDI.Checked = False
                Me._enableAutoInstall = True

            Else
                Me.chkInstallFTDI.Checked = True
                Me._enableAutoInstall = False
            End If

            Me.m5Ints.mdmInst.StartGetListModemInstalledAndFTDIPort(Me.TreeView1, 500, "EstraiFIlesELeggiDevicesInstallati")

            If ClassUser.IsAdministrator Then ReadRegProxySettings(Nothing)
        Catch ex As Exception

        End Try
    End Sub

    'Private usbEvnt As USBEvent
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.m5Ints IsNot Nothing Then
                Me.m5Ints.StopInstall()
            End If
        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "btnCancel", ex, True)
        End Try

    End Sub

    Private Sub btnInstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInstall.Click
        Try
            Me.TabControl1.SelectedTab = TabPage1
            ' Me.m5Ints.StartInstall(Me.chkInstallFTDI.Checked, Me.btnInstall, Me.btnCancel, Me.btnUminstall)
            If ClassEstraiRisorsa.FineEstrazione Then
                Me.m5Ints.StartInstall(True, Me.btnInstall, Me.btnCancel, Me.btnUminstall)
            Else
                MessageBox.Show("Wait file extraction.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If


        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "btnInstall", ex, True)
        End Try
    End Sub

    Public Sub FormMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        Try
            If Me.m5Ints IsNot Nothing AndAlso Me.m5Ints.IsRunning Then

                Select Case MessageBox.Show("Do you want to stop the installation?", "Attention!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                    Case Windows.Forms.DialogResult.Cancel
                        e.Cancel = True

                    Case Windows.Forms.DialogResult.Yes
                        Me.m5Ints.StopInstall()
                        'Me.usbEvnt.Stop()

                    Case Windows.Forms.DialogResult.No
                        e.Cancel = True
                    Case Else

                End Select

            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "FormClosing", ex, True)
        End Try
    End Sub

    Private Sub m5Ints_Finish() Handles m5Ints.InstallFinish
        Try
            Me.m5Ints.mdmInst.GetListModemInstalledAndFTDIPort(Me.TreeView1, 500, "m5Ints_Finish")

            System.Threading.Thread.Sleep(2000)

            Me.m5Ints.RasBook.CreatePhoneBook(Me.m5Ints.mdmInst.listMdmInstalledName)
            Me.m5Ints.RasBook.PrintPhoneBook(Me.TreeView2, Me.m5Ints.mdmInst.listMdmInstalledName)

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "m5Ints_Finish", ex, True)
        End Try
    End Sub

    Private Sub m5Ints_TreeviewPortUpdated(ByVal nPortInstalled As Integer) Handles m5Ints.TreeviewPortUpdated
        Try

            Me.m5Ints.RasBook.PrintPhoneBook(Me.TreeView2, Me.m5Ints.mdmInst.listMdmInstalledName)

            'If ClassFtdiRegKey.FtdiDriverInstalled Then
            '    Safe.CheckBoxSet(False, Me.chkInstallFTDI)
            'Else
            '    Safe.CheckBoxSet(True, Me.chkInstallFTDI)
            'End If

            If Not Me.m5Ints.IsRunning Then AutoIntallModem()

            Me._enableAutoInstall = False

            'If nPortInstalled > 0 Then
            '    Safe.CheckBoxSet(False, Me.CheckBox1)
            'Else
            '    Safe.CheckBoxSet(True, Me.CheckBox1)
            'End If

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "m5Ints_TreeviewPortUpdated", ex, True)
        End Try
    End Sub

    Public Property EnableAutoInstall() As Boolean
        Get
            Return Me._enableAutoInstall
        End Get
        Set(ByVal value As Boolean)
            Me._enableAutoInstall = value
        End Set
    End Property

    Private Sub AutoIntallModem()
        Try

            If Me._enableAutoInstall = False Then
                Exit Sub
            End If

            ' autoinstall
            'Dim s As String = ""
            'For i As Integer = 0 To Me.m5Ints.mdmInst.listPortFTDI_6015Installed.Count - 1
            '    s = s + (i + 1).ToString + " " + Me.m5Ints.mdmInst.listPortFTDI_6015Installed(i) + vbCrLf
            'Next

            'For i As Integer = 0 To Me.m5Ints.mdmInst.listMdmInstalledName.Count - 1
            '    s = s + (i + 1).ToString + " " + Me.m5Ints.mdmInst.listMdmInstalledName(i).Name + "->" + Me.m5Ints.mdmInst.listMdmInstalledName(i).ComPort + vbCrLf
            'Next

            'For i As Integer = 0 To Me.m5Ints.mdmInst.listPortFTDI_6015Plugged.Count - 1
            '    s = s + (i + 1).ToString + " " + Me.m5Ints.mdmInst.listPortFTDI_6015Plugged(i) + vbCrLf
            'Next
            'MessageBox.Show(s)


            Dim portToInstall As String = ""

            ' se c'è almeno una porta plugged
            If ClassMdmInstall.classPortFTDI.GetNumPortsPlugged(Me.m5Ints.mdmInst.listPortFTDI) > 0 Then
                ' trova una porta che NON ha il modem installato
                For i As Integer = 0 To ClassMdmInstall.classPortFTDI.GetNumPortsPlugged(Me.m5Ints.mdmInst.listPortFTDI) - 1
                    If Not IsPortToInstall(Me.m5Ints.mdmInst.listPortFTDI(i).COM.Trim.ToUpper, Me.m5Ints.mdmInst.listMdmInstalledName) Then
                        portToInstall = Me.m5Ints.mdmInst.listPortFTDI(i).COM.Trim.ToUpper
                        Exit For
                    End If
                Next


            End If

            If portToInstall <> "" Then
                '  TopMostMessageBox.ShowNoBlock("Modem on " + portToInstall + " will be installed!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' Me.TabControl1.SelectedTab = TabPage1

                If ClassEstraiRisorsa.WaitFinishExtrcting(10) Then
                    Me.m5Ints.StartInstall(Me.chkInstallFTDI.Checked, Me.btnInstall, Me.btnCancel, Me.btnUminstall)
                Else
                    MessageBox.Show("Wait file extraction.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "AutoIntallModem", ex, True)
        End Try

    End Sub

    Private Function IsPortToInstall(ByVal port As String, ByVal listMdmInstalledName As List(Of LibRasBook.ClassRasBook.classMdmInstalled)) As Boolean
        Try
            For i As Integer = 0 To listMdmInstalledName.Count - 1
                If listMdmInstalledName(i).ComPort.Trim.ToUpper = port Then
                    Return True
                End If
            Next

            Return False

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "IsPortToInstall", ex, True)
        End Try
    End Function

    Private Sub m5Ints_UninstallFinish() Handles m5Ints.UninstallFinish
        Try
            Safe.Control.Visible(False, Me.TableLayoutPanel3)
            Me.m5Ints.mdmInst.GetListModemInstalledAndFTDIPort(Me.TreeView1, 500, "m5Ints_UninstallFinish")

            Me.m5Ints.RasBook.CreatePhoneBook(Me.m5Ints.mdmInst.listMdmInstalledName)
            Me.m5Ints.RasBook.PrintPhoneBook(Me.TreeView2, Me.m5Ints.mdmInst.listMdmInstalledName)

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "m5Ints_UninstallFinish", ex, True)
        Finally
            Safe.Control.Enabled(True, Me.btnUminstall)
        End Try
    End Sub

    Private Sub btnUminstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUminstall.Click
        Try

            Dim msg As String = "Do you want to uninstall " + Me.lastNode.Parent.Text + vbNewLine + vbNewLine + Me.lastNode.Text + " ?"
            Dim port As String = Me.lastNode.Tag.ToString
            Safe.Control.Enabled(False, Me.btnUminstall)

            Dim tipo As ClassM5Installer.TipoUnistall
            If Me.lastNode.Parent.Text.StartsWith("Port") Then
                tipo = ClassM5Installer.TipoUnistall.Port
            Else
                tipo = ClassM5Installer.TipoUnistall.Modem
            End If

            If MessageBox.Show(msg, "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Me.m5Ints.StartUnistallModem(port, tipo, Me.btnInstall, Me.btnUminstall)
            Else
                Safe.Control.Enabled(True, Me.btnUminstall)
                Safe.Control.Visible(False, Me.TableLayoutPanel3)
            End If

        Catch ex As Exception
            ClassMsgErr.LogErr("M5 Setup", "btnUminstall", ex, True)
        Finally

        End Try
    End Sub

    'Protected Overrides Sub WndProc(ByRef m As Message)

    '    MyBase.WndProc(m)

    '    Try
    '        If m.Msg = UsbNotification.WM_DEVICECHANGE Then

    '            'Me.m5Ints.LogAppend("UsbNotification: " + m.Msg.ToString + " " + m.WParam.ToString + " " + m.LParam.ToString + vbCrLf)
    '            Dim devType As Integer = Marshal.ReadInt32(m.LParam, 4)
    '            Dim p As UsbNotification.DEV_BROADCAST_PORT

    '            Select Case CInt(m.WParam)

    '                Case UsbNotification.DBT_DEVICEREMOVECOMPLETE
    '                    If devType = UsbNotification.DBT_DEVTYP_PORT Then
    '                        p = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(UsbNotification.DEV_BROADCAST_PORT)), UsbNotification.DEV_BROADCAST_PORT)
    '                        Usb_DeviceRemoved(p.dbcp_name)
    '                    End If

    '                    ' this is where you do your magic
    '                    Exit Select
    '                Case UsbNotification.DBT_DEVICEARRIVAL
    '                    If devType = UsbNotification.DBT_DEVTYP_PORT Then
    '                        p = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(UsbNotification.DEV_BROADCAST_PORT)), UsbNotification.DEV_BROADCAST_PORT)
    '                        Usb_DeviceAdded(p.dbcp_name)
    '                    End If

    '                    ' this is where you do your magic
    '                    Exit Select
    '                    ' Case Else
    '                    'MessageBox.Show("?: " + m.Msg.ToString)
    '            End Select
    '        End If

    '    Catch ex As Exception
    '    End Try

    'End Sub

    Public Sub Usb_DeviceRemoved(ByVal s As String)
        Try
            If Me.m5Ints.IsRunning = False Then
                Me._enableAutoInstall = False
                Me.m5Ints.LogAppend("Usb port unplugged: " + s + vbCrLf)
                Me.m5Ints.mdmInst.StartGetListModemInstalledAndFTDIPort(Me.TreeView1, 2000, "Usb_DeviceRemoved")
            Else
                Me.m5Ints.LogAppend("Usb port unplugged: " + s + vbCrLf)
            End If
        Catch ex As Exception

        End Try


    End Sub

    Public Sub Usb_DeviceAdded(ByVal s As String)
        Try
            If Me.m5Ints.IsRunning = False Then
                Me._enableAutoInstall = False
                Me.m5Ints.LogAppend("Usb port plugged: " + s + vbCrLf)
                Me.m5Ints.mdmInst.StartGetListModemInstalledAndFTDIPort(Me.TreeView1, 2000, "Usb_DeviceAdded")
            Else
                Me.m5Ints.LogAppend("Usb port plugged: " + s + vbCrLf)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Usb_Generic(ByVal s As String)
        Me.m5Ints.LogAppend("Usb device: " + s + vbCrLf)
    End Sub

    Private lastNode As TreeNode = Nothing

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Try
            ' Dim t As TreeNode = Me.TreeView1.GetNodeAt(e.Location)
            Dim t As TreeNode = e.Node
            If t IsNot Nothing Then
                If t.Level = 2 Then
                    Safe.Control.Visible(True, Me.TableLayoutPanel3)
                    Me.lastNode = t
                    If t.Parent.Text.StartsWith("Port") Then
                        Me.Label2.Text = "Unistall Port:"
                    Else
                        Me.Label2.Text = "Unistall Modem:"
                    End If
                Else
                    Safe.Control.Visible(False, Me.TableLayoutPanel3)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TreeView1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TreeView1.VisibleChanged
        Safe.Control.Visible(False, Me.TableLayoutPanel3)
    End Sub

    Private Sub btnCreateRas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateRas.Click
        Try
            Me.btnCreateRas.Enabled = False

            Me.m5Ints.RasBook.CreatePhoneBook(Me.m5Ints.mdmInst.listMdmInstalledName)
            Me.m5Ints.RasBook.PrintPhoneBook(Me.TreeView2, Me.m5Ints.mdmInst.listMdmInstalledName)
            LibProxy.ClassMcpProxy.ProxyCreate(LibRasBook.ClassRasBook.CreaPathPhonebook)
            ReadRegProxySettings(Nothing)
        Catch ex As Exception
        Finally
            Me.btnCreateRas.Enabled = True
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            My.Computer.Clipboard.SetText(LibRasBook.ClassRasBook.CreaPathPhonebook)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Button2.Enabled = False
            Process.Start(LibRasBook.ClassRasBook.CreaPathPhonebook)
        Catch ex As Exception
        Finally
            Me.Button2.Enabled = True
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            Dim lpszPhonebook As String = "C:\ProgramData\M5\M5.pbk"
            Dim lpszEntry As String = "M5_1"
            'Dim lpRasEntry As New LibRasBook.ClassRasDll.RASENTRY
            'Dim dwEntryInfoSize As Integer = Marshal.SizeOf(lpRasEntry)
            'Dim lpbDeviceInfo As IntPtr
            'Dim dwDeviceInfoSize As Integer

            '' Dim lpbEntry As IntPtr
            'Dim lpdwEntrySize As UInt32
            'Dim lpb As IntPtr
            'Dim lpdwSize As IntPtr

            'lpRasEntry.dwSize = Marshal.SizeOf(lpRasEntry)


            Dim dwEntryInfoSize As UInteger = 0
            Dim i As UInteger = LibRasBook.ClassRasEntryProperties.RasGetEntryProperties(Nothing, "", IntPtr.Zero, dwEntryInfoSize, IntPtr.Zero, IntPtr.Zero)


            Dim getRasEntry As New LibRasBook.ClassRasEntryProperties.RASENTRY
            Dim entryName As String = "M5_1"
            getRasEntry.dwSize = CInt(dwEntryInfoSize)

            Dim ptrRasEntry As IntPtr = Marshal.AllocHGlobal(CInt(dwEntryInfoSize))
            Marshal.StructureToPtr(getRasEntry, ptrRasEntry, False)

            Dim j As UInteger = LibRasBook.ClassRasEntryProperties.RasGetEntryProperties(lpszPhonebook, entryName, ptrRasEntry, dwEntryInfoSize, IntPtr.Zero, IntPtr.Zero)

            Dim outRasEntry As LibRasBook.ClassRasEntryProperties.RASENTRY = DirectCast(Marshal.PtrToStructure(ptrRasEntry, GetType(LibRasBook.ClassRasEntryProperties.RASENTRY)), LibRasBook.ClassRasEntryProperties.RASENTRY)
            Marshal.FreeHGlobal(ptrRasEntry)

            'Dim i As UInteger = LibRasBook.ClassRasEntryProperties.RasGetEntryProperties(Nothing, "", IntPtr.Zero, dwEntryInfoSize, IntPtr.Zero, IntPtr.Zero)



            'Dim esito1 As UInteger = LibRasBook.ClassRasDll.RasGetEntryProperties(lpszPhonebook, lpszEntry, lpRasEntry, lpdwEntrySize, lpb, lpdwSize)

            'Dim esito2 As UInteger = LibRasBook.ClassRasDll.RasSetEntryProperties(lpszPhonebook, lpszEntry, lpRasEntry, dwEntryInfoSize, lpbDeviceInfo, dwDeviceInfoSize)
            'If esito1 = 0 Then

            'End If
            'If esito2 = 0 Then

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private listProxyConn As New List(Of String)
    '  Private Sub ReadRegProxySettings()
    '    System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _readRegProxySettingsTh)
    'End Sub
    Private Sub ReadRegProxySettings(ByVal o As Object)
        Try
            Dim proxyInternetConn As New ProxyInternetSetting
            proxyInternetConn.Get()


            Safe.Control.Text(proxyInternetConn.ProxyAddress, txtAddrProxy)

            Safe.CheckBoxSet(proxyInternetConn.ProxyEnabled, Me.chkProxyEnabled)

            Safe.CheckBoxSet(proxyInternetConn.ProxyBypass, Me.chkProxyBypass)

            Me.listProxyConn = ClassMcpProxy.GetM5Connections()

            Safe.Control.Text("", Me.txtConnections)

            Dim desc As String = ""

            For i As Integer = 0 To listProxyConn.Count - 1
                Dim proxyConn As New ProxyConnections(listProxyConn(i))

                proxyConn.Get()
                desc &= listProxyConn(i) + ": " + proxyConn.ProxyAddress + ", Enabled=" + proxyConn.ProxyEnabled.ToString + ", Bypass=" + proxyConn.ProxyBypass.ToString + vbCrLf


            Next

            Safe.Control.Text(desc, Me.txtConnections)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "ReadRegProxySettings", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Me.m5Ints.RasBook.GetProperties()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If Me.CheckBox1.IsHandleCreated AndAlso Me.CheckBox2.IsHandleCreated AndAlso ClassUser.IsAdministrator Then
                Me.CheckBox1.Enabled = False
                Me.CheckBox2.Enabled = False
                Me.Button5.Enabled = False

                For i As Integer = 0 To listProxyConn.Count - 1
                    Dim proxyConn As New ProxyConnections(listProxyConn(i))
                    proxyConn.Get()
                    proxyConn.ProxyEnabled = CheckBox1.Checked
                    proxyConn.ProxyBypass = CheckBox2.Checked
                    proxyConn.Put()
                Next

                ReadRegProxySettings(Nothing)
            End If
        Catch ex As Exception
        Finally
            Me.CheckBox1.Enabled = True
            Me.CheckBox2.Enabled = True
            Me.Button5.Enabled = True
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Try
            Safe.Control.Visible(False, btnRefresh)
            Me._enableAutoInstall = False
            Me.m5Ints.mdmInst.StartGetListModemInstalledAndFTDIPort(Me.TreeView1, 500, "btnRefresh_Click")
            ' Me.m5Ints.LogAppend("Refresh ..." + vbCrLf)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnUnistallPid6015_Click(sender As Object, e As EventArgs) Handles btnUnistallPid6015.Click
        UnistallPid("6015", btnUnistallPid6015)
    End Sub

    Private Sub btnUnistallPid7018_Click(sender As Object, e As EventArgs) Handles btnUnistallPid7018.Click
        UnistallPid("7018", btnUnistallPid7018)
    End Sub

    Private Sub btnUnistallPid7019_Click(sender As Object, e As EventArgs) Handles btnUnistallPid7019.Click
        UnistallPid("7019", btnUnistallPid7019)
    End Sub

    Private Sub UnistallPid(Pid As String, btn As Button)
        Try

            btn.Enabled = False
            Dim s As String = "Would you like to unistall USB devices and driver" + vbCrLf + "with VID_0403 and PID_" + Pid + "?" + vbCrLf
            If TopMostMessageBox.Show(s, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                Me.txtLogUnistall.Text = "" : Me.txtLogUnistall.Update()

                ClassEstraiRisorsa.EstraiFileFTDIUnistall()

                ClassUnistallFtdi.txtLogUnistall = Me.txtLogUnistall

                ClassUnistallFtdi.Unistall(Pid)
            End If
        Catch ex As Exception
        Finally

            btn.Enabled = True
        End Try
    End Sub

    Private Sub btnDriverList_Click(sender As Object, e As EventArgs) Handles btnDriverList.Click
        Try
            btnDriverList.Enabled = False

            Me.txtLogUnistall.Text = "" : Me.txtLogUnistall.Update()

            ClassEstraiRisorsa.EstraiFileFTDIUnistall()

            ClassUnistallFtdi.txtLogUnistall = Me.txtLogUnistall

            ClassUnistallFtdi.M5DriverList()

        Catch ex As Exception
        Finally
            btnDriverList.Enabled = True
        End Try
    End Sub

    Private Sub lblWin_Click(sender As Object, e As EventArgs) Handles lblWin.Click

    End Sub
End Class
