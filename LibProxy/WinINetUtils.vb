Imports System.Collections.Generic
Imports System.Text

Imports System.Runtime.InteropServices


Friend Class WinINetUtils

#Region "WININET Options"
    Private Const INTERNET_PER_CONN_PROXY_SERVER As UInteger = 2
    Private Const INTERNET_PER_CONN_PROXY_BYPASS As UInteger = 3
    Private Const INTERNET_PER_CONN_FLAGS As UInteger = 1

    Private Const INTERNET_OPTION_REFRESH As UInteger = 37
    Private Const INTERNET_OPTION_PROXY As UInteger = 38
    Private Const INTERNET_OPTION_SETTINGS_CHANGED As UInteger = 39
    Private Const INTERNET_OPTION_END_BROWSER_SESSION As UInteger = 42
    Private Const INTERNET_OPTION_PER_CONNECTION_OPTION As UInteger = 75

    Private Const PROXY_TYPE_DIRECT As UInteger = &H1
    Private Const PROXY_TYPE_PROXY As UInteger = &H2

    Private Const INTERNET_OPEN_TYPE_PROXY As UInteger = 3
#End Region

#Region "STRUCT"
    Private Structure Value1
        Private dwValue As UInteger
        Private pszValue As String
        Private ftValue As FILETIME
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure INTERNET_PER_CONN_OPTION
        Private dwOption As UInteger
        Private Value As Value1
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure INTERNET_PER_CONN_OPTION_LIST
        Private dwSize As UInteger
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=256)> _
        Private pszConnection As String
        Private dwOptionCount As UInteger
        Private dwOptionError As UInteger
        Private pOptions As IntPtr

    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure INTERNET_CONNECTED_INFO
        Private dwConnectedState As Integer
        Private dwFlags As Integer
    End Structure
#End Region

#Region "Interop"
    <DllImport("wininet.dll", EntryPoint:="InternetSetOptionA", CharSet:=CharSet.Ansi, SetLastError:=True, PreserveSig:=True)> _
    Private Shared Function InternetSetOption(ByVal hInternet As IntPtr, ByVal dwOption As UInteger, ByVal pBuffer As IntPtr, ByVal dwReserved As Integer) As Boolean
    End Function
#End Region

    Friend Sub New()
    End Sub


    Friend Shared Sub Notify_OptionSettingChanges()
        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0)
        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0)
    End Sub

End Class
