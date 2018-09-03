
Imports System.Reflection

Public Class FormM5Install

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If LibM5DriverInstall.ClassUser.IsAdministrator Then
            '    Me.Text &= " - " + LibM5DriverInstall.ClassUser.GetName + " (Administrator)" + " " + My.Application.Info.Version.ToString
        Else
            '    Me.Text &= " " + LibM5DriverInstall.ClassUser.GetName + " " + My.Application.Info.Version.ToString
            TopMostMessageBox.ShowNoBlock("You must execute as administrator!", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.M5_Setup1.FormMain_FormClosing(sender, e)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)

        MyBase.WndProc(m)

        Try
            If m.Msg = UsbNotification.WM_DEVICECHANGE Then

                'Me.m5Ints.LogAppend("UsbNotification: " + m.Msg.ToString + " " + m.WParam.ToString + " " + m.LParam.ToString + vbCrLf)
                Dim devType As Integer = System.Runtime.InteropServices.Marshal.ReadInt32(m.LParam, 4)
                Dim p As UsbNotification.DEV_BROADCAST_PORT

                Select Case CInt(m.WParam)

                    Case UsbNotification.DBT_DEVICEREMOVECOMPLETE
                        If devType = UsbNotification.DBT_DEVTYP_PORT Then
                            p = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(m.LParam, GetType(UsbNotification.DEV_BROADCAST_PORT)), UsbNotification.DEV_BROADCAST_PORT)
                            Me.M5_Setup1.Usb_DeviceRemoved(p.dbcp_name)
                        End If

                        ' this is where you do your magic
                        Exit Select
                    Case UsbNotification.DBT_DEVICEARRIVAL
                        If devType = UsbNotification.DBT_DEVTYP_PORT Then
                            p = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(m.LParam, GetType(UsbNotification.DEV_BROADCAST_PORT)), UsbNotification.DEV_BROADCAST_PORT)
                            Me.M5_Setup1.Usb_DeviceAdded(p.dbcp_name)
                        End If

                        ' this is where you do your magic
                        Exit Select
                        ' Case Else
                        'MessageBox.Show("?: " + m.Msg.ToString)
                    Case Else
                        Me.M5_Setup1.Usb_Generic("0x" + (CInt(m.WParam).ToString("X")))
                End Select
            End If

        Catch ex As Exception
        End Try

    End Sub


    Private Sub M5_Setup1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles M5_Setup1.Load

    End Sub
End Class