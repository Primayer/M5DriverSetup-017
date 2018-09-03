Imports System.Collections.Generic
Imports System.Text


Public Class CEngine
    Private _regUtils As New RegUtils()

    Public Sub New()
    End Sub

    Public Sub SetProxyName(ByVal ProxyAddress As String)

        Dim szRegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
        Dim szName As String = "ProxyServer"

        _regUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, szName, ProxyAddress)

    End Sub

    Public Sub EnableProxy(ByVal ProxyAddress As String)

        Dim szRegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
        Dim szName As String = "ProxyEnable"
        Dim szValue As String = String.Empty

        _regUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, "ProxyServer", szValue)

        If szValue <> ProxyAddress Then
            SetProxyName(ProxyAddress)
        End If

        _regUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, szName, 1)

        Dim abValue As Byte()
        Dim aaValue As Char()
        szRegKey = szRegKey & Convert.ToString("Connections")
        _regUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, "DefaultConnectionSettings", abValue)
        aaValue = New Char(abValue.Length - 1) {}

        For i As Integer = 0 To abValue.Length - 1
            If i = 8 Then
                abValue(i) = 3
            End If
            aaValue(i) = Chr(abValue(i))
        Next

        _regUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, "DefaultConnectionSettings", Encoding.ASCII.GetBytes((New String(aaValue)).Replace(szValue, ProxyAddress)))

    End Sub

    Public Sub DisableProxy(ByVal ProxyAddress As String)
        Dim szRegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
        Dim szName As String = "ProxyEnable"
        _regUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, szName, 0)

        Dim abValue As Byte()
        szRegKey = szRegKey & Convert.ToString("Connections")
        _regUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, "DefaultConnectionSettings", abValue)

        For i As Integer = 0 To abValue.Length - 1
            If i = 8 Then
                abValue(i) = 0
                Exit For
            End If
        Next

        _regUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, "DefaultConnectionSettings", abValue)

    End Sub

    Public Function GetProxyName() As String
        Dim szRegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
        Dim szName As String = "ProxyServer"
        Dim szProxyAddress As String = String.Empty

        _regUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, szName, szProxyAddress)

        Return szProxyAddress
    End Function

    Public Function GetProxyStatus() As Integer
        Dim szRegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
        Dim szName As String = "ProxyEnable"
        Dim iProxyStatus As Integer = 0

        _regUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, szRegKey, szName, iProxyStatus)

        Return iProxyStatus
    End Function
End Class

Public MustInherit Class ProxySetting
    Private m_regUtils As New RegUtils()
    Friend ReadOnly Property RegUtils() As RegUtils
        Get
            Return m_regUtils
        End Get
    End Property

    Protected m_proxyAddress As String = ""
    Protected m_proxyEnabled As Boolean = False
    Protected m_proxyBypass As Boolean = False

    Public Property ProxyAddress() As String
        Get
            Return m_proxyAddress
        End Get
        Set(ByVal value As String)
            m_proxyAddress = value
        End Set
    End Property

    Public Property ProxyEnabled() As Boolean
        Get
            Return m_proxyEnabled
        End Get
        Set(ByVal value As Boolean)
            m_proxyEnabled = value
        End Set
    End Property

    Public Property ProxyBypass() As Boolean
        Get
            Return m_proxyBypass
        End Get
        Set(ByVal value As Boolean)
            m_proxyBypass = value
        End Set
    End Property

    Public MustOverride Sub [Get]()
    Public MustOverride Sub Put()
End Class

Public Class ProxyInternetSetting
    Inherits ProxySetting
    Const RegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    Const ProxyKeyName As String = "ProxyServer"
    Const EnabledKeyName As String = "ProxyEnable"
    Const BypassKeyName As String = "ProxyOverride"

    Public Overrides Sub [Get]()
        Try
            RegUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, ProxyKeyName, ProxyAddress)

            Dim enabled As Integer
            RegUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, EnabledKeyName, enabled)
            ProxyEnabled = (enabled = 1)

            Dim bypass As String
            RegUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, BypassKeyName, bypass)
            ProxyBypass = (bypass = "<local>")
        Catch
        End Try
    End Sub

    Public Overrides Sub Put()
        RegUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, ProxyKeyName, ProxyAddress)

        If ProxyEnabled Then
            RegUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, EnabledKeyName, 1)
        Else
            RegUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, EnabledKeyName, 0)
        End If

        If ProxyBypass Then
            RegUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, BypassKeyName, "<local>")
        Else
            RegUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, BypassKeyName, "127.0.0.1")
        End If


    End Sub
End Class

Public Class ProxyConnections
    Inherits ProxySetting
    Const RegKey As String = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
    Private KeyName As String = "DefaultConnectionSettings"
    '
    '            0.  keep this value
    '            1.  "00" placeholder
    '            2.  "00" placeholder
    '            3.  "00" placeholder
    '            4.  "xx" increments if changed
    '            5.  "xx" increments if 4. is "FF"
    '            6.  "00" placeholder
    '            7.  "00" placeholder
    '            8.  "01"=proxy deaktivated; other value=proxy enabled
    '            9.  "00" placeholder
    '            10. "00" placeholder
    '            11. "00" placeholder 
    '            12. "xx" length of "proxyserver:port"
    '            13. "00" placeholder
    '            14. "00" placeholder
    '            15. "00" placeholder
    '            "proxyserver:port"
    '        if 'Bypass proxy for local addresses':::
    '            other stuff with unknown length
    '            "<local>"
    '            36 times "00"
    '        if no 'Bypass proxy for local addresses':::
    '            40 times "00"
    '        


    Public Overrides Sub [Get]()
        Dim ar(-1) As Byte
        RegUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, KeyName, ar)

        If ar.Length < 17 Then
            Exit Sub
        End If

        ProxyEnabled = (ar(8) <> 1)

        Dim i As Integer = 16
        ProxyAddress = ""
        While ar(i) <> 7 AndAlso ar(i) <> 0
            ProxyAddress += System.Convert.ToChar(ar(i)).ToString()
            i += 1
        End While

        Dim bypass As String = ""
        While i < ar.Length
            bypass = ""
            Dim j As Integer = i
            While ar(j) <> 0
                bypass += System.Convert.ToChar(ar(j)).ToString()
                j += 1
            End While
            i += 1
            If bypass.Contains("<local>") Then
                Exit While
            End If
        End While
        ProxyBypass = (bypass.Contains("<local>"))
    End Sub

    Public Overrides Sub Put()
        Dim ar As Byte()
        RegUtils.GetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, KeyName, ar)

        Dim l As New List(Of Byte)()
        For i As Integer = 0 To 15
            If i = 4 Then
                If ar(i) > 0 Then
                    l.Add(0)
                Else
                    l.Add(1)
                End If
            ElseIf i = 5 Then
                l.Add(0)
            ElseIf i = 8 Then
                If ProxyEnabled Then
                    l.Add(CByte(3))
                Else
                    l.Add(CByte(1))
                End If

            ElseIf i = 12 Then
                l.Add(CByte(ProxyAddress.Length))
            Else
                l.Add(ar(i))
            End If
        Next
        For Each c As Char In ProxyAddress
            l.Add(System.Convert.ToByte(c))
        Next
        If ProxyBypass Then
            l.Add(7)
            l.Add(0)
            l.Add(0)
            l.Add(0)
            For Each c As Char In "<local>"
                l.Add(System.Convert.ToByte(c))
            Next
            For i As Integer = 0 To 35
                l.Add(0)
            Next
        Else
            For i As Integer = 0 To 39
                l.Add(0)
            Next
        End If

        RegUtils.SetKeyValue(RegUtils.RegKeyType.CurrentUser, RegKey, KeyName, l.ToArray())
    End Sub

    Public Sub New()

    End Sub
    Public Sub New(ByVal KeyName As String)
        Me.KeyName = KeyName
    End Sub
End Class

Public Class ProxySettings
    Inherits System.Collections.Generic.List(Of ProxySetting)
    Public Property ProxyAddress() As String
        Get
            For Each ps As ProxySetting In Me
                If Not String.IsNullOrEmpty(ps.ProxyAddress) Then
                    Return ps.ProxyAddress
                End If
            Next
            Return ""
        End Get
        Set(ByVal value As String)
            For Each ps As ProxySetting In Me
                ps.ProxyAddress = value
            Next
        End Set
    End Property

    Public Property ProxyEnabled() As Boolean
        Get
            Dim b As Boolean = True
            For Each ps As ProxySetting In Me
                b = b And ps.ProxyEnabled
            Next
            Return b
        End Get
        Set(ByVal value As Boolean)
            For Each ps As ProxySetting In Me
                ps.ProxyEnabled = value
            Next
        End Set
    End Property

    Public Property ProxyBypass() As Boolean
        Get
            Dim b As Boolean = True
            For Each ps As ProxySetting In Me
                b = b And ps.ProxyBypass
            Next
            Return b
        End Get
        Set(ByVal value As Boolean)
            For Each ps As ProxySetting In Me
                ps.ProxyBypass = value
            Next
        End Set
    End Property

    Public Sub [Get]()
        For Each ps As ProxySetting In Me
            ps.[Get]()
        Next
    End Sub

    Public Sub Put()
        For Each ps As ProxySetting In Me
            ps.Put()
        Next
    End Sub

    Public Shared Function GetSimpleAndExtended() As ProxySettings
        Dim psl As New ProxySettings()
        psl.Add(New ProxyInternetSetting())
        psl.Add(New ProxyConnections())

        For i As Integer = 0 To psl.Count - 1
            psl(i).[Get]()
            If String.IsNullOrEmpty(psl(i).ProxyAddress) Then
                psl.RemoveAt(i)
                i -= 1
            End If
        Next

        Return psl
    End Function
End Class
