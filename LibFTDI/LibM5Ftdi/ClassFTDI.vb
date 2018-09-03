Public Class ClassFTDI

    Class ClassUsbM5
        Public Description As String = ""
        Public LocId As Integer = 0
        Public ID As UInteger = 0
        Public ComPort As String = ""
        Public eeX As New FTD2XX_NET.FTDI.FT_XSERIES_EEPROM_STRUCTURE
        Public PID As String = ""
        Public VID As String = ""
        Public Sub New()

        End Sub
    End Class

    Public Function Setrts() As Boolean
        Try

        Catch ex As Exception

        End Try
    End Function

    Public Function SearchM5(ByRef dtms As String) As List(Of ClassUsbM5)

        Dim _sw As New Stopwatch
        Dim _swTot As New Stopwatch
        Try
            _swTot.Reset()
            _swTot.Start()
            _sw.Reset()
            _sw.Start()
            Dim FTDI As New FTD2XX_NET.FTDI
            Dim FTDIstatus As FTD2XX_NET.FTDI.FT_STATUS ' = FTD2XX_NET.FTDI.FT_STATUS.FT_OK
            Dim NumberOfDevices As UInteger = 0
            Dim DeviceList() As FTD2XX_NET.FTDI.FT_DEVICE_INFO_NODE

            Dim M5list As New List(Of ClassUsbM5)

            FTDIstatus = FTDI.GetNumberOfDevices(NumberOfDevices)

            dtms &= "[nDev]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()

            If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then

                If NumberOfDevices = 0 Then
                    Return M5list
                Else

                    ReDim DeviceList(0 To CInt((NumberOfDevices - 1)))
                    FTDIstatus = FTDI.GetDeviceList(DeviceList)

                    dtms &= "[DevList]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()

                    If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then
                        For j As UInteger = 0 To NumberOfDevices - 1UI
                            'cerco i chip FT Device X Series
                            If DeviceList(CInt(j)).Type = FTD2XX_NET.FTDI.FT_DEVICE.FT_DEVICE_X_SERIES Then



                                'mi connetto al dispositivo
                                FTDIstatus = FTDI.OpenByIndex(j)
                                dtms &= "[O." + (j + 1).ToString + "]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()
                                If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then

                                    Dim comport As String = ""

                                    If DeviceList(CInt(j)).Description.ToUpper.IndexOf("HEMINA") >= 0 Then

                                        'identifico il tipo di hardware
                                        FTDIstatus = FTDI.GetCOMPort(comport)


                                        Dim eeX As New FTD2XX_NET.FTDI.FT_XSERIES_EEPROM_STRUCTURE

                                        FTDIstatus = FTDI.ReadXSeriesEEPROM(eeX)
                                        dtms &= "[EE." + (j + 1).ToString + "]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()

                                        Dim p As New ClassUsbM5
                                        p.ComPort = comport
                                        p.Description = DeviceList(CInt(j)).Description
                                        p.ID = DeviceList(CInt(j)).ID
                                        p.LocId = CInt(DeviceList(CInt(j)).LocId)
                                        p.eeX = eeX
                                        p.VID = eeX.VendorID.ToString("X4")
                                        p.PID = eeX.ProductID.ToString("X4")


                                        M5list.Add(p)


                                    End If


                                End If
                                FTDI.Close()
                            End If
                        Next

                    End If
                End If
            End If


            Return M5list

        Catch ex As Exception
        Finally
            dtms &= "[C]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. "
            dtms &= "[Tot]:" + Format(_swTot.Elapsed.TotalSeconds, "0.000") + " sec"
        End Try
    End Function

    Public Function ChangePID(ByRef dtms As String) As List(Of ClassUsbM5)

        Dim _sw As New Stopwatch
        Dim _swTot As New Stopwatch
        Try
            _swTot.Reset()
            _swTot.Start()
            _sw.Reset()
            _sw.Start()
            Dim FTDI As New FTD2XX_NET.FTDI
            Dim FTDIstatus As FTD2XX_NET.FTDI.FT_STATUS ' = FTD2XX_NET.FTDI.FT_STATUS.FT_OK
            Dim NumberOfDevices As UInteger = 0
            Dim DeviceList() As FTD2XX_NET.FTDI.FT_DEVICE_INFO_NODE

            Dim M5list As New List(Of ClassUsbM5)

            FTDIstatus = FTDI.GetNumberOfDevices(NumberOfDevices)

            dtms &= "[nDev]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()

            If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then

                If NumberOfDevices = 0 Then
                    Return M5list
                Else

                    ReDim DeviceList(0 To CInt((NumberOfDevices - 1)))
                    FTDIstatus = FTDI.GetDeviceList(DeviceList)

                    dtms &= "[DevList]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()

                    If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then

                        For j As UInteger = 0 To NumberOfDevices - 1UI

                            'cerco i chip FT Device X Series
                            If DeviceList(CInt(j)).Type = FTD2XX_NET.FTDI.FT_DEVICE.FT_DEVICE_X_SERIES Then

                                'mi connetto al dispositivo
                                FTDIstatus = FTDI.OpenByIndex(j)
                                dtms &= "[O." + (j + 1).ToString + "]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()
                                If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then

                                    Dim comport As String = ""

                                    If DeviceList(CInt(j)).Description.ToUpper.IndexOf("HEMINA") >= 0 Then

                                        'identifico il tipo di hardware
                                        FTDIstatus = FTDI.GetCOMPort(comport)


                                        Dim eeX As New FTD2XX_NET.FTDI.FT_XSERIES_EEPROM_STRUCTURE

                                        FTDIstatus = FTDI.ReadXSeriesEEPROM(eeX)
                                        dtms &= "[EE." + (j + 1).ToString + "]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. " : _sw.Reset() : _sw.Start()

                                        Dim p As New ClassUsbM5
                                        p.ComPort = comport
                                        p.Description = DeviceList(CInt(j)).Description
                                        p.ID = DeviceList(CInt(j)).ID
                                        p.LocId = CInt(DeviceList(CInt(j)).LocId)
                                        p.eeX = eeX
                                        p.VID = eeX.VendorID.ToString("X4")
                                        p.PID = eeX.ProductID.ToString("X4")

                                        If p.Description.StartsWith("HEMINA") AndAlso p.VID = "0403" AndAlso p.PID = "6015" Then

                                            Dim s() As String
                                            s = p.Description.Split("-"c)
                                            If s IsNot Nothing AndAlso s.Length >= 3 Then
                                                Select Case s(1)
                                                    Case "MV255"
                                                        eeX.ProductID = Convert.ToUInt16("7019", 16)
                                                    Case Else
                                                        eeX.ProductID = Convert.ToUInt16("7018", 16)

                                                End Select


                                                p.PID = eeX.ProductID.ToString("X4")

                                                FTDIstatus = FTDI.WriteXSeriesEEPROM(eeX)


                                                '   If FTDIstatus = FTD2XX_NET.FTDI.FT_STATUS.FT_OK Then
                                                M5list.Add(p)
                                                ' End If
                                            End If
                                        End If
                                    End If
                                End If
                                FTDI.Close()
                            End If
                        Next

                    End If
                End If
            End If


            Return M5list

        Catch ex As Exception
        Finally
            dtms &= "[C]:" + Format(_sw.Elapsed.TotalMilliseconds, "0") + " ms. "
            dtms &= "[Tot]:" + Format(_swTot.Elapsed.TotalSeconds, "0.000") + " sec"
        End Try
    End Function


End Class
