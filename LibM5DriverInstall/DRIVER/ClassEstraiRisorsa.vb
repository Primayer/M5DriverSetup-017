

Imports System.IO
Imports System.Reflection

Public Class ClassEstraiRisorsa
    Public Const LibNameSpace As String = "LibM5DriverInstall"

    'Public Const FileFTDIM5_LS As String = "M5Driver7018_LS.zip"
    'Public Const FileFTDIM5_HS As String = "M5Driver7019_HS.zip"
    'Public Const FileFTDIM5 As String = "M5Driver6015.zip"
    Public Const FileFTDIM5 As String = "M5Driver_6015_7018_7019.zip"

    Public Const FileM5_LS_bus_inf As String = "M5_LS_bus.inf"
    Public Const FileM5_LS_port_inf As String = "M5_LS_port.inf"

    Public Const FileM5_HS_bus_inf As String = "M5_HS_bus.inf"
    Public Const FileM5_HS_port_inf As String = "M5_HS_port.inf"


    Public Const FileMdmInf_LS As String = "M5-MDM-LS.inf"
    Public Const FileMdmCat_LS As String = "m5-mdm-ls.cat"
    Public Const FileMdmInf_HS As String = "M5-MDM-HS.inf"
    Public Const FileMdmCat_HS As String = "m5-mdm-hs.cat"

    Public Const FileFTDIUnistall As String = "CmdUninstallFtdi.exe"

    Public Const FileM5UsbList As String = "M5UsbList.exe"
    Public Const FileUsbLib As String = "UsbLib.dll"

    Public Shared FineEstrazione As Boolean = False

    Private Shared lblDescExtractor As Label

    Public Shared Function WaitFinishExtrcting(timeoutsec As Double) As Boolean
        Try
            Dim _Sw As New Stopwatch
            _Sw.Reset()
            _Sw.Start()

            Do
                If FineEstrazione Then
                    Return True
                End If
                If _Sw.Elapsed.TotalSeconds >= timeoutsec Then
                    Return False
                End If
                System.Threading.Thread.Sleep(100)
            Loop
        Catch ex As Exception

        End Try
    End Function



    Public Shared Sub EstraiTutto(lblDescExt As Label)
        Dim f As String = ""
        Try
            lblDescExtractor = lblDescExt

            Dim _sw As New Stopwatch
            _sw.Reset()
            _sw.Start()

            ' Cartella M5Driver:
            Dim c As String = LibPaths.ClassPATH.GetPathM5Driver

            ' Se la cartella esiste già, la cancella con tutto il contenuto
            If My.Computer.FileSystem.DirectoryExists(c) Then
                For Each ftd As String In My.Computer.FileSystem.GetFiles(c, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
                    My.Computer.FileSystem.DeleteFile(ftd)
                Next
            End If

            ' Se la cartella non esiste , la crea
            If Not My.Computer.FileSystem.DirectoryExists(c) Then
                Safe.Control.Text("Create dir", lblDescExtractor)
                My.Computer.FileSystem.CreateDirectory(c)
            End If

            Dim ExecutingAssembly As Assembly = Assembly.GetExecutingAssembly()

            ' Estrae file zip M5Driver_6015_7018_7019.zip da assembly
            f = My.Computer.FileSystem.CombinePath(c, FileFTDIM5)
            WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileFTDIM5, f)

            ' scompatta file zip M5Driver_6015_7018_7019.zip
            EstraiFileZip(c, f)

            ' Estrae files modem da assembly:
            f = My.Computer.FileSystem.CombinePath(c, FileMdmInf_LS)
            WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileMdmInf_LS, f)

            f = My.Computer.FileSystem.CombinePath(c, FileMdmCat_LS)
            WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileMdmCat_LS, f)

            f = My.Computer.FileSystem.CombinePath(c, FileMdmInf_HS)
            WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileMdmInf_HS, f)

            f = My.Computer.FileSystem.CombinePath(c, FileMdmCat_HS)
            WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileMdmCat_HS, f)


            'f = My.Computer.FileSystem.CombinePath(c, FileFTDIM5_LS)
            'WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileFTDIM5_LS, f)
            'EstraiFileZip(c, f)

            'f = My.Computer.FileSystem.CombinePath(c, FileFTDIM5_HS)
            'WriteResourceToFile(ExecutingAssembly, LibNameSpace + "." + FileFTDIM5_HS, f)
            'EstraiFileZip(c, f)

            ' Estrai file M5USBlist
            EstraiFilesM5UsbList()

            _sw.Stop()

        Catch ex As Exception
            Safe.Control.Text("Error extracting files!", lblDescExtractor)
            MessageBox.Show(f + vbCrLf + ex.Message + vbCrLf + ex.StackTrace, "Extract Resource ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Safe.Control.Text("", lblDescExtractor)
            FineEstrazione = True
        End Try
    End Sub

    Private Shared Function EstraiFileZip(cartella As String, nomefileZip As String) As Boolean
        Try
            Dim dc As String = nomefileZip.ToUpper.Replace(".ZIP", "")
            If My.Computer.FileSystem.DirectoryExists(dc) Then
                My.Computer.FileSystem.DeleteDirectory(dc, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If

            'Using zip As New Ionic.Zip.ZipFile(nomefileZip)
            '    zip.ExtractAll(dc, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
            'End Using

            My.Computer.FileSystem.DeleteFile(nomefileZip)

            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Shared Function EstraiFileFTDIUnistall() As Boolean
        Dim f As String = ""

        Try

            Dim a As Assembly = Assembly.GetExecutingAssembly()

            ' Estrai file FTDIUnistall
            Dim c As String = LibPaths.ClassPATH.GetPathM5Driver

            If Not My.Computer.FileSystem.DirectoryExists(c) Then
                My.Computer.FileSystem.CreateDirectory(c)
            End If

            f = My.Computer.FileSystem.CombinePath(c, FileFTDIUnistall)
            WriteResourceToFile(a, LibNameSpace + "." + FileFTDIUnistall, f)

            Return True

        Catch ex As Exception
            MessageBox.Show(f + vbCrLf + ex.Message, "Extract Resource ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Shared Function EstraiFilesM5UsbList() As Boolean
        Dim f As String = ""

        Try
            Dim c As String = ""

            Dim a As Assembly = Assembly.GetExecutingAssembly()

            ' Estrai file M5USBlist
            c = LibPaths.ClassPATH.GetPathM5USBList
            If Not My.Computer.FileSystem.DirectoryExists(c) Then
                My.Computer.FileSystem.CreateDirectory(c)
            End If

            f = My.Computer.FileSystem.CombinePath(c, FileM5UsbList)
            WriteResourceToFile(a, LibNameSpace + "." + FileM5UsbList, f)

            f = My.Computer.FileSystem.CombinePath(c, FileUsbLib)
            WriteResourceToFile(a, LibNameSpace + "." + FileUsbLib, f)

            Return True

        Catch ex As Exception
            MessageBox.Show(f + vbCrLf + ex.Message, "Extract Resource ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


    Public Shared Sub WriteResourceToFile(ByVal targetAssembly As Assembly, ByVal resourceName As String, ByVal filepath As String)
        Try
            'Dim ar() As String
            'ar = targetAssembly.GetManifestResourceNames()

            'Dim k As String = ""
            'If ar IsNot Nothing Then
            '    k = "Cerca: " + resourceName + " in " + vbCrLf + targetAssembly.FullName + vbCrLf + vbCrLf
            '    For i As Integer = 0 To ar.Length - 1
            '        k = k + ar(i) + vbCrLf
            '    Next
            '    k = k + vbCrLf
            '    k = k + "Load: " + targetAssembly.GetName().Name & "." & resourceName
            '    MessageBox.Show(k)
            'End If

            'Using s As Stream = targetAssembly.GetManifestResourceStream(targetAssembly.GetName().Name & "." & resourceName)

            Safe.Control.Text(My.Computer.FileSystem.GetName(filepath), lblDescExtractor)

            Using s As Stream = targetAssembly.GetManifestResourceStream(resourceName)
                If s Is Nothing Then
                    Throw New Exception("Cannot find embedded resource '" & resourceName & "'")
                End If

                Dim buffer(0 To CInt(s.Length) - 1) As Byte
                s.Read(buffer, 0, buffer.Length)

                Dim c As String = My.Computer.FileSystem.GetParentPath(filepath)
                If Not My.Computer.FileSystem.DirectoryExists(c) Then
                    My.Computer.FileSystem.CreateDirectory(c)
                End If
                Using sw As BinaryWriter = New BinaryWriter(File.Open(filepath, FileMode.Create))
                    sw.Write(buffer)
                End Using
            End Using

        Catch ex As Exception

            MessageBox.Show(ex.Message, "WriteResourceToFile", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Safe.Control.Text("", lblDescExtractor)
        End Try

    End Sub
End Class
