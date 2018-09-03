Imports System.Collections.Generic
Imports System.Text

Imports System.Runtime.InteropServices


Friend Class User32Utils
#Region "USER32 Options"
    Shared HWND_BROADCAST As New IntPtr(&HFFFF)
    Shared WM_SETTINGCHANGE As New IntPtr(&H1A)
#End Region

#Region "STRUCT"
    Private Enum SendMessageTimeoutFlags As UInteger
        SMTO_NORMAL = &H0
        SMTO_BLOCK = &H1
        SMTO_ABORTIFHUNG = &H2
        SMTO_NOTIMEOUTIFNOTHUNG = &H8
    End Enum
#End Region

#Region "Interop"
    '[DllImport("user32.dll", CharSet = CharSet.Auto)]
    'public static extern int SendMessage(int hWnd, int msg, int wParam, IntPtr lParam);

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function SendMessageTimeout(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As UIntPtr, ByVal lParam As UIntPtr, ByVal fuFlags As SendMessageTimeoutFlags, ByVal uTimeout As UInteger, _
     ByRef lpdwResult As UIntPtr) As IntPtr
    End Function
#End Region


    Friend Sub New()
    End Sub

    Friend Shared Sub Notify_SettingChange()
        Dim result As UIntPtr
        SendMessageTimeout(HWND_BROADCAST, CUInt(WM_SETTINGCHANGE), UIntPtr.Zero, UIntPtr.Zero, SendMessageTimeoutFlags.SMTO_NORMAL, 1000, _
         result)
    End Sub

End Class
