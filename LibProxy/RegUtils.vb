Imports System.Collections.Generic
Imports System.Text

Imports Microsoft.Win32


Public Class RegUtils
    Public Enum RegKeyType As Integer
        CurrentUser = 1
        LocalMachine = 2
    End Enum

    Public Sub New()
    End Sub

    ' byte[]
    Public Sub GetKeyValue(ByVal KeyType As RegKeyType, ByVal RegKey As String, ByVal Name As String, ByRef Value As Byte())
        Dim oRegKey As RegistryKey = Nothing

        Select Case CInt(KeyType)
            Case 1
                oRegKey = Registry.CurrentUser
                Exit Select
            Case 2
                oRegKey = Registry.LocalMachine
                Exit Select
        End Select
        oRegKey = oRegKey.OpenSubKey(RegKey)

        Value = DirectCast(oRegKey.GetValue(Name), Byte())

        oRegKey.Close()
    End Sub

    ' int
    Public Sub GetKeyValue(ByVal KeyType As RegKeyType, ByVal RegKey As String, ByVal Name As String, ByRef Value As Integer)
        Dim oRegKey As RegistryKey = Nothing

        Select Case CInt(KeyType)
            Case 1
                oRegKey = Registry.CurrentUser
                Exit Select
            Case 2
                oRegKey = Registry.LocalMachine
                Exit Select
        End Select
        oRegKey = oRegKey.OpenSubKey(RegKey)

        Value = CInt(oRegKey.GetValue(Name, Nothing))

        oRegKey.Close()

    End Sub

    ' string
    Public Sub GetKeyValue(ByVal KeyType As RegKeyType, ByVal RegKey As String, ByVal Name As String, ByRef Value As String)
        Dim oRegKey As RegistryKey = Nothing

        Select Case CInt(KeyType)
            Case 1
                oRegKey = Registry.CurrentUser
                Exit Select
            Case 2
                oRegKey = Registry.LocalMachine
                Exit Select
        End Select
        oRegKey = oRegKey.OpenSubKey(RegKey)

        Value = DirectCast(oRegKey.GetValue(Name), String)

        oRegKey.Close()

    End Sub

    ' byte[]
    Friend Sub SetKeyValue(ByVal KeyType As RegKeyType, ByVal RegKey As String, ByVal Name As String, ByVal Value As Byte())
        Dim oRegKey As RegistryKey = Nothing

        Select Case CInt(KeyType)
            Case 1
                oRegKey = Registry.CurrentUser
                Exit Select
            Case 2
                oRegKey = Registry.LocalMachine
                Exit Select
        End Select

        oRegKey = oRegKey.OpenSubKey(RegKey, True)
        oRegKey.SetValue(Name, Value)
        oRegKey.Close()
        User32Utils.Notify_SettingChange()
        WinINetUtils.Notify_OptionSettingChanges()
    End Sub

    ' string
    Friend Sub SetKeyValue(ByVal KeyType As RegKeyType, ByVal RegKey As String, ByVal Name As String, ByVal Value As String)
        Dim oRegKey As RegistryKey = Nothing

        Select Case CInt(KeyType)
            Case 1
                oRegKey = Registry.CurrentUser
                Exit Select
            Case 2
                oRegKey = Registry.LocalMachine
                Exit Select
        End Select

        oRegKey = oRegKey.OpenSubKey(RegKey, True)
        oRegKey.SetValue(Name, Value)
        oRegKey.Close()
        User32Utils.Notify_SettingChange()
        WinINetUtils.Notify_OptionSettingChanges()
    End Sub

    ' int
    Friend Sub SetKeyValue(ByVal KeyType As RegKeyType, ByVal RegKey As String, ByVal Name As String, ByVal Value As Integer)
        Dim oRegKey As RegistryKey = Nothing

        Select Case CInt(KeyType)
            Case 1
                oRegKey = Registry.CurrentUser
                Exit Select
            Case 2
                oRegKey = Registry.LocalMachine
                Exit Select
        End Select

        oRegKey = oRegKey.OpenSubKey(RegKey, True)
        oRegKey.SetValue(Name, Value)
        oRegKey.Close()
        User32Utils.Notify_SettingChange()
        WinINetUtils.Notify_OptionSettingChanges()
    End Sub
End Class
