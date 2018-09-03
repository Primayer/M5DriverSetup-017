Public Class ClassSerachHeminaUSBraw


    '24/05/2018 12:12				
    '2,366 sec.  Found 2 devices				
    'HEMINA-MV110-HW00	 USB M5 (COM14)	 COM14	 USB\VID_0403&PID_6015\7&3695DA3A&0&1	 Port_#0001.Hub_#0002
    'HEMINA-MV252-HW01	 USB M5 (COM15)	 COM15	 USB\VID_0403&PID_6015\7&3695DA3A&0&3	 Port_#0003.Hub_#0002

    Public Function FullPath_M5UsbListexe() As String
        Try
            Return My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5USBList, ClassEstraiRisorsa.FileM5UsbList)
        Catch ex As Exception

        End Try
    End Function

    Public Function NomeFileToSave() As String
        Try
            Return My.Computer.FileSystem.CombinePath(LibPaths.ClassPATH.GetPathM5USBList, "M5UsbList.txt")
        Catch ex As Exception

        End Try
    End Function

    Class classUSB
        Public Data As String = ""
        Public TimeToRead As String = ""
        Public NumDevices As String = ""
        Public DeviceList As New List(Of classUSBraw)
        Public FileContent As String = ""
        Public FilePath As String = ""
    End Class

    Class classUSBraw
        Implements ICloneable

        Public USBraw As String = ""
        Public USBdescriprion As String = ""
        Public COM As String = ""
        Public USBVidPisSn As String = ""
        Public Loacaion As String = ""
        Public VID As String = ""
        Public PID As String = ""
        Public M5Speed As String = ""

        Public Function Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone
        End Function
    End Class

    Public Function SearchHeminaUSB(FullPath_M5UsbListexe As String, FullPath_Save As String, ByRef Usb As classUSB) As Boolean
        Try
            Usb = New classUSB

            If Not My.Computer.FileSystem.FileExists(FullPath_M5UsbListexe) Then
                If ClassEstraiRisorsa.EstraiFilesM5UsbList() = False Then
                    Return False
                End If
            End If

            ' prepara processo
            Dim pi As New ProcessStartInfo
            pi.CreateNoWindow = True
            pi.WindowStyle = ProcessWindowStyle.Hidden
            pi.FileName = FullPath_M5UsbListexe
            pi.Arguments = FullPath_Save

            ' lancia processo
            Dim p As Process = Process.Start(pi)

            ' attende che finisca
            p.WaitForExit()

            ' carica file con dati salvati
            Return LoadLastUSBfileSaved(FullPath_Save, Usb)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "SearchHeminaUSB", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    Public Function LoadLastUSBfileSaved(nomeFile As String, ByRef Usb As classUSB) As Boolean
        Try
            Dim r As String = ""
            Dim l As New List(Of String)
            Dim lUsb As New List(Of classUSBraw)
            Usb.FilePath = nomeFile
            Using sr As New System.IO.StreamReader(nomeFile)
                Do While Not sr.EndOfStream
                    r = sr.ReadLine
                    Usb.FileContent &= r + vbCrLf
                    If r IsNot Nothing Then
                        l.Add(r.Trim.TrimEnd(";"c))
                    End If
                Loop

            End Using

            For i As Integer = 0 To l.Count - 1
                Dim s() As String
                s = l(i).Split(";"c)
                If s IsNot Nothing Then
                    If i = 0 Then
                        Usb.Data = s(0)
                    ElseIf i = 1 Then
                        If s.Length > 0 Then
                            Dim c() As String
                            c = s(0).Split("."c)
                            If c IsNot Nothing Then
                                If c.Length > 0 Then Usb.TimeToRead = c(0)
                                If c.Length > 1 Then Usb.NumDevices = c(1)
                            End If
                        End If
                    Else
                        Dim u As New classUSBraw
                        'HEMINA-MV252-HW01	 USB M5 (COM15)	 COM15	 USB\VID_0403&PID_6015\7&3695DA3A&0&3	 Port_#0003.Hub_#0002

                        If s.Length >= 1 Then u.USBraw = s(0).Trim
                        If s.Length >= 2 Then u.USBdescriprion = s(1).Trim
                        If s.Length >= 3 Then u.COM = s(2).Trim
                        If s.Length >= 4 Then u.USBVidPisSn = s(3).Trim
                        If s.Length >= 5 Then u.Loacaion = s(4).Trim


                        Dim f() As String
                        f = u.USBVidPisSn.Split("\"c)
                        If f IsNot Nothing AndAlso f.Length >= 2 Then
                            Dim ID As String = f(1)
                            If ID IsNot Nothing AndAlso ID.Trim <> "" AndAlso ID.IndexOf("&") >= 0 Then
                                ID = ID.Replace("&"c, "+"c)
                            End If
                            ClassFtdiRegKey.EstraiVidPid(ID, ClassFtdiRegKey.EnumVidPidSeparator.Più, u.VID, u.PID)
                        End If

                        If s.Length >= 8 Then u.M5Speed = s(7).Trim

                        lUsb.Add(u)

                    End If

                End If
            Next

            Usb.DeviceList = lUsb

            Return True

        Catch ex As Exception
            ' MessageBox.Show(ex.Message, "LoadLastUSBfileSaved", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
End Class
