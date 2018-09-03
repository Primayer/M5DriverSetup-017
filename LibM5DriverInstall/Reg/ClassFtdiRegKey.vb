Imports Microsoft.Win32

Public Class ClassFtdiRegKey

    Public Const FtdiVidPid7018 As String = "VID_0403+PID_7018"
    Public Const FtdiVidPid7019 As String = "VID_0403+PID_7019"
    Public Const FtdiVidPid6015 As String = "VID_0403+PID_6015"

    Public Const keyFTDIBUS As String = "SYSTEM\CurrentControlSet\Enum\FTDIBUS\"
    Public Const subKeyNameDeviceParameters As String = "Device Parameters"
    Public Const KeyPortName As String = "PortName"
    Public Const KeyConfigData As String = "ConfigData"
    Public Const keyEmulationMode As String = "EmulationMode"
    Public Const keyEnum As String = "SYSTEM\CurrentControlSet\Enum\"
    Public Const Ftdi As String = "FTDIBUS"

    Class ClassFTDIInstalled
        Public PortName As String = ""
        Public ConfigData As String = ""
        Public PID As String = ""
        Public VID As String = ""
    End Class

    ''' <summary>
    ''' Search PID_7018, PID_7019, PID_6015 with M5_SPEED field, in SYSTEM\CurrentControlSet\Enum\FTDIBUS\
    ''' </summary>
    ''' <param name="errore">Error</param>
    ''' <returns></returns>
    Public Shared Function SearchFtdiInstalled(ByRef errore As Boolean) As List(Of ClassFTDIInstalled)

        Dim listPortCom As New List(Of ClassFTDIInstalled)

        Try
            errore = False

            Dim kfti As RegistryKey = Nothing


            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            ' apre "SYSTEM\CurrentControlSet\Enum\FTDIBUS\"
            kfti = rk.OpenSubKey(keyFTDIBUS)

            If kfti IsNot Nothing AndAlso kfti.GetSubKeyNames IsNot Nothing AndAlso kfti.GetSubKeyNames.Length > 0 Then

                For Each s As String In kfti.GetSubKeyNames
                    'VID_0403+PID_6015
                    'VID_0403+PID_7018
                    'VID_0403+PID_7018+6&adae852&0&3
                    'VID_0403+PID_7019

                    If s.ToUpper.StartsWith(FtdiVidPid7018.ToUpper) Or
                       s.ToUpper.StartsWith(FtdiVidPid7019.ToUpper) Or
                       s.ToUpper.StartsWith(FtdiVidPid6015.ToUpper) Then

                        For Each md As String In rk.OpenSubKey(keyFTDIBUS + "\" + s).GetSubKeyNames
                            ' 0000
                            For Each v As String In rk.OpenSubKey(keyFTDIBUS + "\" + s + "\" + md).GetSubKeyNames

                                If v IsNot Nothing AndAlso v.ToUpper = subKeyNameDeviceParameters.ToUpper Then
                                    'Device Parameters
                                    Dim val As String = rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters).GetValue(KeyPortName).ToString
                                    Dim cd As Byte() = CType(rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters).GetValue(KeyConfigData), Byte())

                                    If val IsNot Nothing AndAlso val <> "" Then
                                        Dim p As New ClassFTDIInstalled
                                        p.PortName = val
                                        EstraiVidPid(s.Replace("&"c, "+"c), EnumVidPidSeparator.Più, p.VID, p.PID)

                                        Dim buff As String = ""
                                        For i As Integer = 0 To cd.Length - 1
                                            buff = buff + "[" + cd(i).ToString("X2") + "]"
                                        Next
                                        p.ConfigData = buff

                                        Dim ns() As String = rk.OpenSubKey(keyFTDIBUS + "\" + s + "\" + md + "\" + subKeyNameDeviceParameters).GetValueNames
                                        Dim M5Speed As String = ""

                                        If ReadM5_Speed(ns, rk, keyFTDIBUS, s, md, M5Speed) Then

                                            listPortCom.Add(p)

                                            'AN232B-05 Configuring FT232R, FT2232 And FT232B Baud Rates
                                            'http://www.ftdichip.com/Support/Documents/AppNotes/AN232B-05_BaudRates.pdf

                                            'USB Transfer Size: trasmit and receive: 3F,3F (index 2 e 3)
                                            '  0 is 64 bytes, and 3F is (63+1)*64 = 4096.

                                            'FtdiPort232.NT.HW.AddReg]

                                            'HKR,,"ConfigData",1, ...
                                            '11,00,3F,3F,  10,27,00,00,  88,13,00,00,  C4,09,00,00,  E2,04,00,00,  71,02,00,00,  38,41,00,00,  9C,80,00,00,  
                                            '4E,C0,00,00,  34,00,00,00,  1A,00,00,00,  0D,00,00,00,  06,40,00,00,  03,80,00,00,  00,00,00,00,  D0,80,00,00.

                                            ' indici:
                                            ' 0, 1, 2, 3,   4, 5, 6, 7,   8, 9,10,11,  12,13,14,15,  16,17,18,19,  20,21,22,23,  24,25,26,27,  28,29,30,31,  
                                            '32,33,34,35,  36,37,38,39,  40,41,42,43,  44,45,46,47,  48,49,50,51,  52,53,54,55,  56,57,58,59,  69,61,62,63. 

                                            ' .....          300             600          1200          2400           4800           9600        19200       
                                            '  38400        57600         115200        230400        460800         921600         RESERVED      14400

                                            '[01][02]= 11,00
                                            '[03][04]= 3F,3F                                                               2    7    1    0
                                            '[04][05]= 10,27 => divisor = 10000, rate = 300    --->  10,27 --->  0x2710 = 0010 0111 0001 0000 b = 10000 d; rate: 3E6/10000=300
                                            '[06][07]= 00,00                                                                                   (i bit 15,14 sono per il sub int div: 01b=0.5). rate: 3E6/312.5=9600
                                            '[08][09]= 88,13 => divisor = 5000,  rate = 600
                                            '[10][11]= 00,00
                                            '[12][13]= C4,09 => divisor = 2500,  rate = 1200
                                            '[14][15]= 00,00
                                            '[16][17]= E2,04 => divisor = 1250,  rate = 2,400
                                            '[18][19]= 00,00
                                            '[20][21]= 71,02 => divisor = 625,   rate = 4,800
                                            '[22][23]= 00,00                                                               4    1    3    8
                                            '[24][25]= 38,41 => divisor = 312.5, rate = 9,600  --->  38,41 --->  0x4138 = 0100 0001 0011 1000 b = 16669 d; int div: 16669 and 0x3fff=0x9C=312d
                                            '[26][27]= 00,00                                                               8    0    9    C    (i bit 15,14 sono per il sub int div: 01b=0.5). rate: 3E6/312.5=9600
                                            '[28][29]= 9C,80 => divisor = 156,   rate = 19,230 --->  9C,80 --->  0x809C = 1000 0000 1001 1100 b = 32924 d; int div: 32924 and 0x3fff=0x9C=156d
                                            '[30][31]= 00,00                                                              ..-- ---- ---- ----  (i bit 15,14 sono per il sub int div: 10b=0.25). rate: 3E6/156.25=19200 
                                            '[32][33]= 4E,C0 => divisor = 78,    rate = 38,461       questi sopra sono valori di default!
                                            '[34][35]= 00,00
                                            '[36][37]= 34,00 => divisor = 52,    rate = 57,692
                                            '[38][39]= 00,00
                                            '[40][41]= 1A,00 => divisor = 26,    rate = 115,384
                                            '[42][43]= 00,00
                                            '[44][45]= 0D,00 => divisor = 13,    rate = 230,769
                                            '[46][47]= 00,00
                                            '[48][49]= 06,40 => divisor = 6.5,   rate = 461,538
                                            '[50][51]= 00,00
                                            '[52][53]= 03,80 => divisor = 3.25,  rate = 923,076
                                            '[54][55]= 00,00
                                            '[56][57]= 00,00 => RESERVED
                                            '[58][59]= 00,00
                                            '[60][61]= D0,80 => divisor = 208.25,rate= 14406
                                            '[62][63]= 00,00

                                            ' bit per sub int div
                                            '15,14 = 00 - sub-integer divisor = 0
                                            '15,14 = 01 - sub-integer divisor = 0.5
                                            '15,14 = 10 - sub-integer divisor = 0.25
                                            '15,14 = 11 - sub-integer divisor = 0.125

                                            'M5  funziona con un baudrate a 1500000 (1 mega e mezzo); 
                                            ' Su windows possiamo scegliere varie velocità...300,600,1200 ecc.
                                            ' se scegliamo che ai 1200 di windows corrispondano 1.5mega allora: modifichiamo i bytes della velocità 1200, cioè i bytes [12][13] da "C4,09" a "02,00"
                                            '                                 0    0    0    2
                                            ' e quindi: 02,00 --->  0x0002 = 0000 0000 0000 0010 b; sub int div: 00b=0; rate: 3E6/2.0=1.5E6

                                            ' Win baud 1200-> M5 a 1.5mega; questa impostazione non serve
                                            ' cd(12) = 2
                                            ' cd(13) = 0

                                            ' Win baud 19200-> M5 a 1.5mega (default quando si installa  M2700 con mdmhayes.inf
                                            ''''''''''''''' tolta il 03/05/2018 perchè inclus nel driver M5'''''''''''''''''''''''''''''''''''''''
                                            'cd(28) = 2
                                            'cd(29) = 0
                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                            ' Esempio di come impostare un baud rate di 19200
                                            ' 1) dagli indici di 19200 cioè [28],[29]  ottengo 9C,80 (vedi sopra)
                                            ' 2) giro i bytes: 9C,80 -> 80,9C=32924d=1000000010011100
                                            '                             indici bit:5432109876543210
                                            '             sub-integer divisor:15,14 :--..............= 10 -> 0.25  (vedi  bit per sub int div)
                                            '                         divisore:13:0 :..--------------= 00000010011100=0x9C=156d
                                            ' 3) 3000000/156.25 = 19200 baud, cioè alla velocità 19200 ho imposto un divisore che con freq 19200
                                            ' se voglio far corrispondere ai 19200 ina velocità di 1.5M devo mettere negli indici [28],[29]
                                            '   3000000/1.5M= 3000000/1500000=2.0  cioè bit 15,14=0,0 e bit 13:0=00000000000010=0x2
                                            ' poi giro i bytes 02,00  ->negli indici [28],[29] 

                                            'A divisor of 0 will give 3 MBaud. Sub-integer divisors between 0 And 2 are Not allowed.

                                            ' Pertarnto ci saranno 2 driver per il modem, uno con velocità 19200 M5-MDM-LS.inf e uno con velocità 38400 M5-MDM-HS.inf
                                            ' un unico driver Ftdi con alias di velocità a 19200 e 38400
                                            ' 1) per velocità 1.5 Mbit con i bytes 02,00 agli indici [28],[29]->19200 low speed
                                            ' 2) per velocità 3.0 Mbit con i bytes 00,00 agli indici [32],[33]->38400 high speed

                                            'USB Transfer Size:
                                            ' cd(2) = 0 ' size=F(y)=(y+1)*64; default 0x3F=63; se y=0 ->F(63)=(63+1)*64=4096 bytes; se y=0 ->F(0)=(0+1)*64=64 bytes
                                            ' cd(3) = 0 ' size=F(y)=(y+1)*64; default 0x3F=63; se y=0 ->F(63)=(63+1)*64=4096 bytes; se y=0 ->F(0)=(0+1)*64=64 bytes


                                            ''''''''''''''' tolta il 03/05/2018 perchè inclus nel driver M5'''''''''''''''''''''''''''''''''''''''
                                            ' If changeBaudRate Then rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyName, True).SetValue(KeyConfigData, cd, RegistryValueKind.Binary)
                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                        End If
                                    End If
                                End If

                            Next
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            errore = True
            ' ClassMsgErr.LogErr("ClassFtdiRegKey", "SearchFtdiInstalled", ex, True)
        End Try

        Return listPortCom

    End Function

    Public Shared Function SearchFtdi6015_NoM5SpeedInstalled(ByRef errore As Boolean) As List(Of ClassFTDIInstalled)

        Dim listPortCom As New List(Of ClassFTDIInstalled)

        Try
            errore = False

            Dim kfti As RegistryKey = Nothing


            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            ' apre "SYSTEM\CurrentControlSet\Enum\FTDIBUS\"
            kfti = rk.OpenSubKey(keyFTDIBUS)

            If kfti IsNot Nothing AndAlso kfti.GetSubKeyNames IsNot Nothing AndAlso kfti.GetSubKeyNames.Length > 0 Then

                For Each s As String In kfti.GetSubKeyNames
                    'VID_0403+PID_6015
                    'VID_0403+PID_7018
                    'VID_0403+PID_7018+6&adae852&0&3
                    'VID_0403+PID_7019

                    If s.ToUpper.StartsWith(FtdiVidPid6015.ToUpper) Then

                        For Each md As String In rk.OpenSubKey(keyFTDIBUS + "\" + s).GetSubKeyNames
                            ' 0000
                            For Each v As String In rk.OpenSubKey(keyFTDIBUS + "\" + s + "\" + md).GetSubKeyNames

                                If v IsNot Nothing AndAlso v.ToUpper = subKeyNameDeviceParameters.ToUpper Then
                                    'Device Parameters
                                    Dim val As String = rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters).GetValue(KeyPortName).ToString
                                    Dim cd As Byte() = CType(rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters).GetValue(KeyConfigData), Byte())

                                    If val IsNot Nothing AndAlso val <> "" Then
                                        Dim p As New ClassFTDIInstalled
                                        p.PortName = val
                                        EstraiVidPid(s.Replace("&"c, "+"c), EnumVidPidSeparator.Più, p.VID, p.PID)

                                        Dim buff As String = ""
                                        For i As Integer = 0 To cd.Length - 1
                                            buff = buff + "[" + cd(i).ToString("X2") + "]"
                                        Next
                                        p.ConfigData = buff

                                        Dim ns() As String = rk.OpenSubKey(keyFTDIBUS + "\" + s + "\" + md + "\" + subKeyNameDeviceParameters).GetValueNames
                                        Dim M5Speed As String = ""

                                        If Not ReadM5_Speed(ns, rk, keyFTDIBUS, s, md, M5Speed) Then
                                            listPortCom.Add(p)
                                        End If
                                    End If
                                End If

                            Next
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            errore = True
            ' ClassMsgErr.LogErr("ClassFtdiRegKey", "SearchFtdiInstalled", ex, True)
        End Try

        Return listPortCom

    End Function

    Public Shared Function ReadM5_Speed(ns() As String, rk As Microsoft.Win32.RegistryKey, keyFTDIBUS As String, key_xxxx As String, keyS As String, ByRef M5Speed As String) As Boolean
        Try
            M5Speed = ""

            If ns IsNot Nothing AndAlso ns.Length > 0 Then
                Dim idx As Integer = -1
                For i As Integer = 0 To ns.Length - 1
                    If ns(i).ToUpper = "M5_SPEED" Then
                        idx = i
                        Exit For
                    End If
                Next
                If idx >= 0 Then
                    M5Speed = rk.OpenSubKey(keyFTDIBUS + "\" + key_xxxx + "\" + keyS + "\" + "Device Parameters").GetValue("M5_SPEED").ToString
                    Return True
                End If
            End If

        Catch ex As Exception

        End Try
    End Function

    Enum EnumVidPidSeparator
        Più
        Ecomm
    End Enum

    Public Shared Function EstraiVidPid(VIDPID As String, sepType As EnumVidPidSeparator, ByRef VID As String, ByRef PID As String) As Boolean
        Try
            'VID_0403+PID_7018
            VID = ""
            PID = ""
            Const SepP As Char = "+"c
            Const SepE As Char = "&"c
            Dim sep As Char = SepP

            If sepType = EnumVidPidSeparator.Più Then
                sep = SepP
            End If
            If sepType = EnumVidPidSeparator.Ecomm Then
                sep = SepE
            End If

            VIDPID = VIDPID.ToUpper

            If VIDPID.IndexOf("VID_") >= 0 AndAlso VIDPID.IndexOf("PID_") >= 0 AndAlso VIDPID.IndexOf(sep) >= 0 Then

                VIDPID = VIDPID.Replace("VID_", "")
                VIDPID = VIDPID.Replace("PID_", "")

                '0403+7018

                Dim s() As String
                s = VIDPID.Split(sep)

                If s IsNot Nothing AndAlso s.Length >= 2 Then
                    VID = s(0)
                    PID = s(1)
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If

        Catch ex As Exception

        End Try
    End Function

    Class ClassDeviceID
        Public PortName As String = ""
        Public Indice As String = ""
        Public DeviceID As String = ""
        Public ID As String = ""
    End Class


    Public Shared Function FtdiDriverInstalled() As Boolean
        Try



            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            For Each s As String In rk.OpenSubKey(keyEnum).GetSubKeyNames
                If s.ToUpper.StartsWith(Ftdi.ToUpper) Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            '  ClassMsgErr.LogErr("ClassFtdiRegKey", "FtdiDriverInstalled", ex, True)
        End Try
    End Function

    Public Shared Function SearchFtdiDeviceID() As List(Of ClassDeviceID)
        'ricerca  : HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\FTDIBUS\VID_0403+PID_6015......\0000
        '           HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\FTDIBUS\VID_0403+PID_6015......\0001



        Dim listDeviceID As New List(Of ClassDeviceID)

        Try



            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine



            For Each s As String In rk.OpenSubKey(keyFTDIBUS).GetSubKeyNames

                If s.ToUpper.StartsWith(FtdiVidPid7018.ToUpper) Or
                   s.ToUpper.StartsWith(FtdiVidPid7019.ToUpper) Or
                   s.ToUpper.StartsWith(FtdiVidPid6015.ToUpper) Then

                    For Each md As String In rk.OpenSubKey(keyFTDIBUS + "\" + s).GetSubKeyNames


                        Dim p As New ClassDeviceID

                        p.Indice = md
                        p.DeviceID = Ftdi + "\" + s + "\" + md
                        p.ID = s
                        Try
                            p.PortName = rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters).GetValue(KeyPortName).ToString
                            listDeviceID.Add(p)
                        Catch ex As Exception

                        End Try

                    Next
                End If
            Next

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassFtdiRegKey", "SearchFtdiDeviceID", ex, True)
        End Try

        Return listDeviceID

    End Function

    Public Shared Function SetEmulationMode(ByVal PortNameToSet As String, ByVal OnOff As Boolean) As Boolean

        Try



            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            For Each s As String In rk.OpenSubKey(keyFTDIBUS).GetSubKeyNames

                If s.ToUpper.StartsWith(FtdiVidPid7018.ToUpper) Then

                    For Each md As String In rk.OpenSubKey(keyFTDIBUS + "\" + s).GetSubKeyNames

                        For Each v As String In rk.OpenSubKey(keyFTDIBUS + "\" + s + "\" + md).GetSubKeyNames

                            If v.ToUpper = subKeyNameDeviceParameters.ToUpper Then
                                Dim valPortName As String = rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters).GetValue(KeyPortName).ToString

                                If valPortName IsNot Nothing AndAlso valPortName.ToUpper = PortNameToSet.ToUpper Then
                                    Dim val As Integer = 0
                                    If OnOff Then
                                        val = 1
                                    End If
                                    rk.OpenSubKey(keyFTDIBUS + s + "\" + md).OpenSubKey(subKeyNameDeviceParameters, True).SetValue(keyEmulationMode, val, RegistryValueKind.DWord)
                                    rk.Flush()
                                    rk.Close()
                                    Return True
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            Return False
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassFtdiRegKey", "SetEmulationMode", ex, True)
        End Try
    End Function

    Public Shared Function ReadUsbFlags(ByVal lstBox As ListBox, ByRef keyReg As String) As Boolean
        Try
            If lstBox IsNot Nothing Then lstBox.Items.Clear()

            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\usbflags

            keyReg = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\usbflags"

            For Each s As String In rk.OpenSubKey("SYSTEM\CurrentControlSet\Control\usbflags").GetValueNames
                If rk.OpenSubKey("SYSTEM\CurrentControlSet\Control\usbflags").GetValueKind(s) = RegistryValueKind.Binary Then
                    Dim val() As Byte = CType(rk.OpenSubKey("SYSTEM\CurrentControlSet\Control\usbflags").GetValue(s), Byte())
                    If lstBox IsNot Nothing Then lstBox.Items.Add(s + "=" + val(0).ToString)
                End If
            Next

            Return True

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassFtdiRegKey", "ReadUsbFlags", ex, True)
        End Try
    End Function

    Public Shared Function WriteUsbFlags(ByVal OnOff As Boolean) As Boolean
        Try
            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

            'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\usbflags

            Dim b() As Byte = {1}

            If OnOff Then
                b(0) = 1
            Else
                b(0) = 0
            End If

            rk.OpenSubKey("SYSTEM\CurrentControlSet\Control\usbflags", True).SetValue("IgnoreHWSerNum04036015", b, RegistryValueKind.Binary)
            rk.OpenSubKey("SYSTEM\CurrentControlSet\Control\usbflags", True).SetValue("IgnoreHWSerNum04037018", b, RegistryValueKind.Binary)
            rk.OpenSubKey("SYSTEM\CurrentControlSet\Control\usbflags", True).SetValue("IgnoreHWSerNum04037019", b, RegistryValueKind.Binary)

        Catch ex As Exception
            ClassMsgErr.LogErr("ClassFtdiRegKey", "WriteUsbFlags", ex, True)
        End Try
    End Function

    Public Shared Function GetListModemInstalled(ByVal lstBox As ListBox, ByVal SetFlowControlToNone As Boolean, ByVal SetDriverDesc As Boolean) As Boolean
        Try
            Const keyModem As String = "SYSTEM\CurrentControlSet\Control\Class\{4D36E96D-E325-11CE-BFC1-08002BE10318}"
            Const keyAttachedTo As String = "AttachedTo"
            Const keyFriendlyName As String = "FriendlyName"
            Const keyProperties As String = "Properties"
            Const keyDriverDesc As String = "DriverDesc"

            If lstBox IsNot Nothing Then lstBox.Items.Clear()

            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim numero As Integer = 0

            For Each key_xxxx As String In rk.OpenSubKey(keyModem).GetSubKeyNames

                If Integer.TryParse(key_xxxx, numero) Then

                    Dim AttachedTo As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyAttachedTo).ToString
                    Dim FriendlyName As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyFriendlyName).ToString

                    If lstBox IsNot Nothing Then lstBox.Items.Add(key_xxxx + "\" + keyAttachedTo + " : " + AttachedTo + " " + FriendlyName)

                    'If md.ToUpper = "DriverDesc".ToUpper Then
                    '    rk.OpenSubKey(keyModem + "\" + s, True).SetValue("DriverDesc", "M5", RegistryValueKind.String)
                    'End If
                    'If md.ToUpper = "FriendlyName".ToUpper Then
                    '    rk.OpenSubKey(keyModem + "\" + s, True).SetValue("FriendlyName", "M5", RegistryValueKind.String)
                    'End If


                    Dim cd As Byte() = CType(rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyProperties), Byte())
                    Dim buff As String = ""
                    For i As Integer = 0 To cd.Length - 1
                        buff = buff + "[" + cd(i).ToString("X2") + "]"
                    Next
                    If lstBox IsNot Nothing Then lstBox.Items.Add(keyProperties + ": " + buff)

                    ' toglie controllo di flusso: (vedi Procalc tools do wdk)
                    ' cd(20) = 32 ' 0x20=32dec
                    cd(20) = 0 ' 0x20=32dec

                    If SetFlowControlToNone Then
                        rk.OpenSubKey(keyModem + "\" + key_xxxx, True).SetValue(keyProperties, cd, RegistryValueKind.Binary)
                        Dim k1 As Microsoft.Win32.RegistryKey
                        k1 = rk.OpenSubKey(keyModem + "\" + key_xxxx, True)
                        k1.SetValue(keyProperties, cd, RegistryValueKind.Binary)
                        k1.Flush()
                        k1.Close()
                    End If

                    If SetDriverDesc Then
                        Dim k2 As Microsoft.Win32.RegistryKey
                        k2 = rk.OpenSubKey(keyModem + "\" + key_xxxx, True)
                        k2.SetValue(keyDriverDesc, "M5", RegistryValueKind.String)
                        k2.Flush()
                        k2.Close()
                    End If
                End If
            Next
        Catch ex As Exception
            ' ClassMsgErr.LogErr("ClassFtdiRegKey", "GetListModemInstalled", ex, True)
        End Try

    End Function

    Private Sub ColoraPortComPlugged(ByVal tree As TreeView, ByVal PortNameValue As String)
        Try
        Catch ex As Exception
        End Try
    End Sub

    Public Shared Function GetModemInstalled(ByVal PortNAme As String, ByVal SetDriverDesc As Boolean) As Boolean
        Try
            Const keyModem As String = "SYSTEM\CurrentControlSet\Control\Class\{4D36E96D-E325-11CE-BFC1-08002BE10318}"
            Const keyAttachedTo As String = "AttachedTo"
            Const keyFriendlyName As String = "FriendlyName"
            Const keyProperties As String = "Properties"
            Const keyDriverDesc As String = "DriverDesc"


            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine


            Dim numero As Integer = 0

            For Each key_xxxx As String In rk.OpenSubKey(keyModem).GetSubKeyNames

                If Integer.TryParse(key_xxxx, numero) Then

                    Dim sb() As String
                    sb = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValueNames
                    If sb IsNot Nothing AndAlso sb.Length > 0 Then
                        Dim trovatoAttachedTo As Boolean = False

                        For i As Integer = 0 To sb.Length - 1
                            If sb(i) = keyAttachedTo Then
                                trovatoAttachedTo = True
                                Exit For
                            End If
                        Next

                        If trovatoAttachedTo Then
                            Dim AttachedTo As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyAttachedTo).ToString
                            'Dim FriendlyName As String = rk.OpenSubKey(keyModem + "\" + key_xxxx).GetValue(keyFriendlyName).ToString

                            If PortNAme.Trim.ToUpper = AttachedTo.Trim.ToUpper Then
                                If SetDriverDesc Then
                                    Dim k2 As Microsoft.Win32.RegistryKey
                                    k2 = rk.OpenSubKey(keyModem + "\" + key_xxxx, True)
                                    k2.SetValue(keyDriverDesc, "M5", RegistryValueKind.String)
                                    k2.Flush()
                                    k2.Close()
                                End If
                                Return True
                            End If
                        End If
                    End If

                End If
            Next
        Catch ex As Exception
            ClassMsgErr.LogErr("ClassFtdiRegKey", "GetModemInstalled", ex, True)
        End Try

    End Function
End Class
