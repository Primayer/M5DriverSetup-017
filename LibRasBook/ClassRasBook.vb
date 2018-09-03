Imports DotRas
Imports System.Collections.ObjectModel
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
 
Public Class ClassRasBook

    ' Const path As String = "C:\Users\msimonato.HEMINA\AppData\Roaming\Microsoft\Network\Connections\Pbk\rasphone.pbk"

    Private _ConnectionOptions As RasConnectionOptions

    Class classMdmInstalled
        Public Name As String = ""
        Public ComPort As String = ""
    End Class

    Public ReadOnly Property ConnectionOptions() As RasConnectionOptions
        Get
            Return _ConnectionOptions
        End Get
    End Property

    Public Function GetEntryNames(ByVal lst As ListBox, ByVal NomeRasPhone As String) As List(Of String)
        Try
            Dim l As New List(Of String)

            If lst IsNot Nothing Then lst.Items.Clear()

            Using pb As New DotRas.RasPhoneBook

                pb.Open(NomeRasPhone)

                For Each pbe As DotRas.RasEntry In pb.Entries
                    If pbe.Device.DeviceType = RasDeviceType.Modem Then
                        l.Add(pbe.Name)
                        If lst IsNot Nothing Then lst.Items.Add(pbe.Name + ": " + pbe.Device.Name + " [" + pbe.Device.DeviceType.ToString + "]")
                    End If

                Next
            End Using
            Return l

        Catch ex As Exception
            MessageBox.Show(ex.Message, "GetEntryNames", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    'Public Sub CreatePhoneBook(ByVal lst As ListBox, ByVal FileRasphoneBook As String, ByVal filtroCavodiCom As String)

    '    Try
    '        Dim n As Integer = 0

    '        Dim l As List(Of String)
    '        l = GetEntryNames(Nothing, FileRasphoneBook)
    '        lst.Items.Clear()

    '        Dim pb As New DotRas.RasPhoneBook
    '        pb.Open(FileRasphoneBook)
    '        pb.Entries.Clear()

    '        ' MessageBox.Show(pb.GetPhoneBookPath(RasPhoneBookType.AllUsers))
    '        For Each dev As RasDevice In RasDevice.GetDevices

    '            If dev.DeviceType = RasDeviceType.Modem AndAlso dev.Name.ToUpper.StartsWith(filtroCavodiCom.ToUpper) Then
    '                n += 1

    '                Dim e As DotRas.RasEntry = DotRas.RasEntry.CreateDialUpEntry("M5_" + n.ToString, "0", dev)
    '                'e.PrerequisitePhoneBook = path
    '                'e.Update()

    '                e.EncryptionType = RasEncryptionType.None

    '                e.FramingProtocol = RasFramingProtocol.Ppp
    '                e.IPv4InterfaceMetric = 0
    '                e.NetworkProtocols.IPv6 = False

    '                'BindMsNetClient=0


    '                e.Options.ModemLights = True
    '                'e.Options.ShowDialingProgress = True
    '                e.Options.SoftwareCompression = False
    '                e.Options.DisableLcpExtensions = True
    '                e.Options.DoNotNegotiateMultilink = True


    '                e.Options.DoNotUseRasCredentials = True
    '                e.Id.ToString()

    '                pb.Entries.Add(e)
    '                lst.Items.Add("Name: """ + dev.Name + """ [type: """ + dev.DeviceType.ToString + """] ")

    '            End If
    '        Next

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "GetDevices", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Function EntryExist(ByVal nameEntry As String, ByVal FileRasphoneBook As String) As RasEntry
        Try

            Dim PhoneBook As New RasPhoneBook
            PhoneBook.Open(FileRasphoneBook)

            For Each ee As RasEntry In PhoneBook.Entries
                If ee.Name = nameEntry Then
                    Return ee

                End If
            Next

            Return Nothing
        Catch ex As Exception

        End Try
    End Function

    Public Sub GetProperties()
        Try
            Dim FileRasphoneBook As String = CreaPathPhonebook()

            Dim dwEntryInfoSize As UInteger = 0
            Dim i As UInteger = ClassRasOpt.RasGetEntryProperties(Nothing, "", IntPtr.Zero, dwEntryInfoSize, IntPtr.Zero, IntPtr.Zero)


            Dim getRasEntry As New ClassRasOpt.RASENTRY()
            Dim entryName As String = "M5_1"
            getRasEntry.dwSize = CInt(dwEntryInfoSize)

            Dim ptrRasEntry As IntPtr = Marshal.AllocHGlobal(CInt(dwEntryInfoSize))
            Marshal.StructureToPtr(getRasEntry, ptrRasEntry, False)

            Dim j As UInteger = ClassRasOpt.RasGetEntryProperties(FileRasphoneBook, entryName, ptrRasEntry, dwEntryInfoSize, IntPtr.Zero, IntPtr.Zero)

            Dim outRasEntry As ClassRasOpt.RASENTRY = DirectCast(Marshal.PtrToStructure(ptrRasEntry, GetType(ClassRasOpt.RASENTRY)), ClassRasOpt.RASENTRY)



            Dim x1 As Boolean = ClassRasOpt.GetPropValue(ClassRasOpt.RASEO.ModemLights, CUInt(outRasEntry.dwfOptions))
            Dim x2 As Boolean = ClassRasOpt.GetProp2Value(ClassRasOpt.RASEO2.SecureClientForMSNet, CUInt(outRasEntry.dwfOptions))
            Dim x3 As Boolean = ClassRasOpt.GetProp2Value(ClassRasOpt.RASEO2.SecureFileAndPrint, CUInt(outRasEntry.dwfOptions))


            Marshal.FreeHGlobal(ptrRasEntry)

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "GetProperties", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub CreatePhoneBook(ByVal listModem As List(Of LibRasBook.ClassRasBook.classMdmInstalled))

        Try
            Dim n As Integer = 0
            Dim FileRasphoneBook As String = CreaPathPhonebook()

            If listModem.Count = 0 Then
                Exit Sub
            End If

            Dim filtroCavodiCom As String = "" 'listModem(0).Name.Split("#"c)(0)

            ' vedi stringa M2700 = "M5-MDM-HS" nel file inf (;Localizable)


            filtroCavodiCom = "M5-"



            filtroCavodiCom = filtroCavodiCom.Trim

            Using pb As New DotRas.RasPhoneBook
                pb.Open(FileRasphoneBook)

                pb.Entries.Clear()

                Dim ldevOk As New List(Of String)

                ' MessageBox.Show(pb.GetPhoneBookPath(RasPhoneBookType.AllUsers))

                Dim ld As New List(Of String)
                For Each dev As RasDevice In RasDevice.GetDevices
                    If dev.DeviceType = RasDeviceType.Modem Then
                        ld.Add(dev.Name + "; " + dev.DeviceType.ToString)
                    End If

                Next
                Dim ldr As System.Collections.ObjectModel.ReadOnlyCollection(Of DotRas.RasDevice)

                ldr = RasDevice.GetDevices

                For Each dev As RasDevice In RasDevice.GetDevices

                    If dev.DeviceType = RasDeviceType.Modem AndAlso dev.Name.ToUpper.StartsWith(filtroCavodiCom.ToUpper) Then
                        n += 1
                        ldevOk.Add(dev.Name)
                        Dim nameEntry As String = "M5_" + n.ToString

                        Dim e As DotRas.RasEntry = EntryExist(nameEntry, FileRasphoneBook)
                        If e Is Nothing Then
                            e = DotRas.RasEntry.CreateDialUpEntry(nameEntry, "0", dev)
                        End If


                        e.FramingProtocol = RasFramingProtocol.Ppp
                        e.EncryptionType = RasEncryptionType.None
                        e.FrameSize = 0

                        e.Options.DisableLcpExtensions = True
                        e.Options.DisableNbtOverIP = True
                        e.Options.DoNotNegotiateMultilink = True
                        e.Options.DoNotUseRasCredentials = True
                        e.Options.Internet = False
                        e.Options.IPHeaderCompression = False
                        e.Options.ModemLights = True
                        e.Options.NetworkLogOn = False

                        e.Options.PreviewDomain = False
                        e.Options.PreviewPhoneNumber = False
                        e.Options.PreviewUserPassword = False
                        e.Options.PromoteAlternates = False

                        e.Options.ReconnectIfDropped = False
                        e.Options.RemoteDefaultGateway = False

                        e.Options.RequireChap = False
                        e.Options.RequireDataEncryption = False
                        e.Options.RequireEap = False
                        e.Options.RequireEncryptedPassword = False
                        e.Options.RequireMSChap = False
                        e.Options.RequireMSChap2 = False
                        e.Options.RequireMSEncryptedPassword = False
                        e.Options.RequirePap = False
                        e.Options.RequireSpap = False
                        e.Options.RequireWin95MSChap = False

                        e.Options.SecureClientForMSNet = True
                        e.Options.SecureFileAndPrint = True
                        e.Options.SecureLocalFiles = True

                        e.Options.SharedPhoneNumbers = False
                        e.Options.SharePhoneNumbers = False
                        e.Options.ShowDialingProgress = False

                        e.Options.SoftwareCompression = False
                        e.Options.TerminalAfterDial = False
                        e.Options.TerminalBeforeDial = False

                        e.Options.UseCountryAndAreaCodes = False
                        e.Options.UseGlobalDeviceSettings = False
                        e.Options.UseLogOnCredentials = False
                        e.Options.UsePreSharedKey = False




                        If ModemIsInlist(dev.Name, listModem) Then
                            If Not pb.Entries.Contains(e.Name) Then
                                pb.Entries.Add(e)
                            End If
                            If e IsNot Nothing Then e.Update()
                        End If



                        'lst.Items.Add("Name: """ + dev.Name + """ [type: """ + dev.DeviceType.ToString + """] ")

                    End If
                Next
                If ldevOk.Count > 0 Then

                End If
            End Using
            'If n = 0 Then
            '    MessageBox.Show("Not found any device installed!", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            'Else
            '    MessageBox.Show("Found " + n.ToString + " device installed!", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            'End If

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "GetDevices", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function RemoveEntryBydevName(ByVal devNameToRemove As String) As Boolean

        Try
            Dim n As Integer = 0
            Dim FileRasphoneBook As String = CreaPathPhonebook()

            Dim ini As New LibRasBook.ClassIni(FileRasphoneBook)

            Dim devName As String = ""

            For Each nameEntry As String In ini.GetEntries
                If ini.GetValue(nameEntry, "Device", devName) Then
                    If devName.ToUpper = devNameToRemove.ToUpper Then
                        Using pb As New DotRas.RasPhoneBook
                            pb.Open(FileRasphoneBook)
                            pb.Entries(nameEntry).Remove()
                        End Using
                        Return True
                    End If
                End If
            Next

            Return False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RemoveEntryBydevName", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Class classdata
        Public tree As TreeView
        Public listModem As List(Of LibRasBook.ClassRasBook.classMdmInstalled)
    End Class

    Public Sub StartPrintPhoneBook(ByVal tree As TreeView, ByVal listModem As List(Of LibRasBook.ClassRasBook.classMdmInstalled))
        Try
            Dim p As New classdata
            p.tree = tree
            p.listModem = listModem
            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _PrintPhoneBook, p)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub PrintPhoneBook(ByVal tree As TreeView, ByVal listModem As List(Of LibRasBook.ClassRasBook.classMdmInstalled))
        Dim p As New classdata
        p.tree = tree
        p.listModem = listModem
        _PrintPhoneBook(p)
    End Sub

    Class classEntry
        Public EntryName As String = ""
        Public ComPort As String = ""
        Public Device As String = ""
    End Class

    Public Function GetFileEntryList(ByVal FileRasphoneBook As String) As List(Of classEntry)
        Dim l As New List(Of classEntry)
        Try


            Dim ini As New LibRasBook.ClassIni(FileRasphoneBook)
            Dim Portvalue As String = ""
            Dim Devicevalue As String = ""

            If Not My.Computer.FileSystem.FileExists(FileRasphoneBook) Then
                Return l
            End If

            For Each nameEntry As String In ini.GetEntries
                If ini.GetValue(nameEntry, "Port", Portvalue) AndAlso ini.GetValue(nameEntry, "PreferredDevice", Devicevalue) Then 'PreferredDevice

                    Dim p As New classEntry

                    p.ComPort = Portvalue
                    p.Device = Devicevalue
                    p.EntryName = nameEntry
                    l.Add(p)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "ClassRasBook.GetEntryList", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return l
    End Function

    Public Function GetEntryList() As List(Of classEntry)
        Dim l As New List(Of classEntry)
        Try

            Dim FileRasphoneBook As String = CreaPathPhonebook()
            Dim ini As New LibRasBook.ClassIni(FileRasphoneBook)
            Dim Portvalue As String = ""
            Dim Devicevalue As String = ""

            If Not My.Computer.FileSystem.FileExists(FileRasphoneBook) Then
                Return l
            End If

            For Each nameEntry As String In ini.GetEntries
                If ini.GetValue(nameEntry, "Port", Portvalue) AndAlso ini.GetValue(nameEntry, "PreferredDevice", Devicevalue) Then 'PreferredDevice

                    Dim p As New classEntry

                    p.ComPort = Portvalue
                    p.Device = Devicevalue
                    p.EntryName = nameEntry
                    l.Add(p)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace, "ClassRasBook.GetEntryList", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return l
    End Function


    Private Sub _PrintPhoneBook(ByVal o As Object)
        Try
            Dim cd As classdata = CType(o, classdata)
            Dim FileRasphoneBook As String = CreaPathPhonebook()
            Dim ini As New LibRasBook.ClassIni(FileRasphoneBook)
            Dim Portvalue As String = ""
            Dim Devicevalue As String = ""

            Safe.TreeView.NodesClear(cd.tree)

            Dim ne As Integer = ini.GetEntries.Count

            For Each nameEntry As String In ini.GetEntries
                If ini.GetValue(nameEntry, "Port", Portvalue) AndAlso ini.GetValue(nameEntry, "PreferredDevice", Devicevalue) Then 'PreferredDevice

                    ' stampo solo se la porta appartiene alla lista di FTDI232
                    If PortIsInlist(Portvalue, cd.listModem) Then
                        Dim p As New Safe.TreeView.ClassTreeData
                        Dim f As New Safe.TreeView.ClassTreeData

                        Safe.TreeView.TreeAddImg(nameEntry, cd.tree, 0, p)
                        Safe.TreeView.NodesAddImg(Portvalue + " <=> " + Devicevalue, p.Nodo, 1, f)
                    Else
                        ' rimuove entry in phonebook
                        Using pb As New DotRas.RasPhoneBook
                            pb.Open(FileRasphoneBook)
                            pb.Entries(nameEntry).Remove()
                        End Using
                    End If

                End If
            Next

            Safe.TreeView.ExpandAll(cd.tree)
        Catch ex As Exception

        End Try
    End Sub

    Private Function PortIsInlist(ByVal port As String, ByVal listModem As List(Of LibRasBook.ClassRasBook.classMdmInstalled)) As Boolean
        Try
            For i As Integer = 0 To listModem.Count - 1
                If listModem(i).ComPort.ToUpper = port.ToUpper Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception

        End Try
    End Function
    Private Function ModemIsInlist(ByVal modem As String, ByVal listModem As List(Of LibRasBook.ClassRasBook.classMdmInstalled)) As Boolean
        Try
            For i As Integer = 0 To listModem.Count - 1
                If listModem(i).Name.ToUpper = modem.ToUpper Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception

        End Try
    End Function
    Public Function CreaDialUpEntry() As Boolean
        Try

        Catch ex As Exception

        End Try
    End Function

    'LcpExtensions=0
    'ShareMsFilePrint=0
    'BindMsNetClient=0
    'NegotiateMultilinkAlways=0

    'ShareMsFilePrint=0
    'BindMsNetClient=0

    Public Shared Function CreaPathPhonebook() As String
        Try
            Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
            '  Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

            path = My.Computer.FileSystem.CombinePath(path, "M5")

            If Not My.Computer.FileSystem.DirectoryExists(path) Then
                My.Computer.FileSystem.CreateDirectory(path)
            End If

            path = My.Computer.FileSystem.CombinePath(path, "M5.pbk")

            Return path

        Catch ex As Exception

        End Try
    End Function


End Class
