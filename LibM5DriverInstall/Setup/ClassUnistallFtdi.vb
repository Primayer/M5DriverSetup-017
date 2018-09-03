Public Class ClassUnistallFtdi

    Public Shared txtLogUnistall As TextBox

    Public Shared Function Unistall(PID As String) As Boolean
        Try

            Dim c As String = LibPaths.ClassPATH.GetPathM5Driver
            Dim FullPath_FileFTDIUnistall As String = My.Computer.FileSystem.CombinePath(c, ClassEstraiRisorsa.FileFTDIUnistall)
            Dim f As String = My.Computer.FileSystem.CombinePath(c, "Log_PID_" + PID + ".txt")

            If txtLogUnistall IsNot Nothing Then
                Safe.Control.Text("Unistalling VID_0403 PID_" + PID + " ...", txtLogUnistall)
            End If


            If Not My.Computer.FileSystem.FileExists(FullPath_FileFTDIUnistall) Then
                If ClassEstraiRisorsa.EstraiFilesM5UsbList() = False Then
                    Return False
                End If
            End If

            ' prepara processo  'https://www.dotnetperls.com/redirectstandardoutput
            Dim pi As New ProcessStartInfo
            pi.CreateNoWindow = True
            pi.WindowStyle = ProcessWindowStyle.Hidden
            pi.UseShellExecute = False
            pi.RedirectStandardOutput = True
            pi.FileName = FullPath_FileFTDIUnistall
            pi.Arguments = "0403 " + PID

            ' lancia processo


            Dim p As New Process

            p.StartInfo.Verb = "runas"

            p = Process.Start(pi)

            ' legge stream di output
            Using rd As New System.IO.StreamReader(p.StandardOutput.BaseStream)
                Dim s As String = rd.ReadToEnd
                If txtLogUnistall IsNot Nothing Then
                    Safe.Control.Text(s, txtLogUnistall)
                End If
            End Using

            ' attende che finisca
            p.WaitForExit()

            Return True

        Catch ex As Exception

        End Try
    End Function

    Class classM5Driver
        Public FileInf As String = ""
        Public Provider As String = ""
        Public Classe As String = ""
        Public Versione As String = ""
        Public Firmatario As String = ""


        '        Nome pubblicato :             oem108.inf
        'Provider di pacchetti driver:   Hemina
        'Classe:                     Controller USB(Universal Serial Bus)
        'Versione e data driver:   07/17/2018 1.1.00
        'Nome firmatario
    End Class

    Public Shared Function M5DriverList() As List(Of classM5Driver)
        Try

            Dim l As New List(Of classM5Driver)

            ' prepara processo
            Dim pi As New ProcessStartInfo
            pi.CreateNoWindow = True
            pi.WindowStyle = ProcessWindowStyle.Hidden

            pi.UseShellExecute = False
            pi.RedirectStandardOutput = True
            pi.FileName = "pnputil.exe"
            pi.Arguments = "-e"

            ' lancia processo  'https://www.dotnetperls.com/redirectstandardoutput
            Dim p As Process = Process.Start(pi)

            ' legge stream di output
            Using rd As New System.IO.StreamReader(p.StandardOutput.BaseStream, System.Text.Encoding.UTF7)

                Dim r1 As String = ""
                Dim r2 As String = ""
                Dim r3 As String = ""
                Dim r4 As String = ""
                Dim r5 As String = ""

                Do

                    If _leggiRigaExit(rd, r1) Then
                        Exit Do
                    End If

                    If r1.ToUpper.IndexOf(".INF") >= 0 Then
                        Dim pinf As New classM5Driver
                        pinf.FileInf = DopoDuePunti(r1)
                        If _leggiRigaExit(rd, r2) Then
                            Exit Do
                        End If
                        pinf.Provider = DopoDuePunti(r2)
                        If r2.ToUpper.IndexOf("HEMINA") >= 0 OrElse r2.ToUpper.IndexOf("FTDI") >= 0 Then

                            If _leggiRigaExit(rd, r3) Then
                                Exit Do
                            End If
                            pinf.Classe = DopoDuePunti(r3)

                            If _leggiRigaExit(rd, r4) Then
                                Exit Do
                            End If
                            pinf.Versione = DopoDuePunti(r4)

                            If _leggiRigaExit(rd, r5) Then
                                Exit Do
                            End If
                            pinf.Firmatario = DopoDuePunti(r5)

                            l.Add(pinf)
                        End If

                    End If
                Loop

            End Using

            For i As Integer = 0 To l.Count - 1
                Safe.Textbox.AppendText((i + 1).ToString + ") " + l(i).FileInf + ": " + l(i).Provider + ", " + l(i).Classe + " [" + l(i).Versione + "]" + vbCrLf, txtLogUnistall)
            Next

            ' attende che finisca
            p.WaitForExit()

            Return l
        Catch ex As Exception

        End Try
    End Function
    Public Shared Function InstallDriverpPnputil(fileinf As String) As Boolean
        Try

            Dim l As New List(Of classM5Driver)

            ' prepara processo
            Dim pi As New ProcessStartInfo
            pi.CreateNoWindow = True
            pi.WindowStyle = ProcessWindowStyle.Hidden

            pi.UseShellExecute = False
            pi.RedirectStandardOutput = True
            pi.FileName = "C:\Windows\System32\pnputil.exe"
            pi.Arguments = "-a """ + fileinf + """"

            ' lancia processo  'https://www.dotnetperls.com/redirectstandardoutput
            Dim p As Process = Process.Start(pi)

            Dim so As String = ""
            ' legge stream di output
            Using rd As New System.IO.StreamReader(p.StandardOutput.BaseStream, System.Text.Encoding.UTF7)

                so = rd.ReadToEnd


            End Using
            Safe.Textbox.AppendText(so + vbCrLf, txtLogUnistall)

            ' attende che finisca
            p.WaitForExit()

            Return True
        Catch ex As Exception

        End Try
    End Function

    Private Shared Function DopoDuePunti(r As String) As String
        Try
            If r.IndexOf(":"c) >= 0 Then
                Dim s() As String
                s = r.Split(":"c)
                If s IsNot Nothing AndAlso s.Length >= 2 Then
                    Return s(1).Replace(vbTab, "").Trim
                End If
            End If
            Return r
        Catch ex As Exception

        End Try
    End Function
    Private Shared Function _leggiRigaExit(rd As System.IO.StreamReader, ByRef r As String) As Boolean
        Try
            r = ""

            Dim riga As String = rd.ReadLine
            If riga Is Nothing Then
                Return True
            Else
                r = riga
                Return False
            End If
        Catch ex As Exception

        End Try
    End Function

End Class
