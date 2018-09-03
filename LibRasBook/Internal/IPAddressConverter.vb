Imports System.ComponentModel
Imports System.Net
Imports System.Net.Sockets

Namespace Ras.Internal

    ''' <summary>
    ''' Provides a type converter used for converting <see cref="System.Net.IPAddress"/> values.
    ''' </summary>
    Friend Class IPAddressConverter
        Inherits TypeConverter
#Region "Constructors"

        ''' <summary>
        ''' Initializes a new instance of the <see cref="IPAddressConverter"/> class.
        ''' </summary>
        Public Sub New()
        End Sub

#End Region

#Region "Methods"

        ''' <summary>
        ''' Returns whether this converter can convert an object of the given type to the type of this converter.
        ''' </summary>
        ''' <param name="context">An <see cref="ITypeDescriptorContext"/> that provides a format context.</param>
        ''' <param name="sourceType">The source type.</param>
        ''' <returns><b>true</b> if the conversion can be performed, otherwise <b>false</b>.</returns>
        '        Public Overrides Function CanConvertFrom(ByVal context As ITypeDescriptorContext, ByVal sourceType As Type) As Boolean
        '            If sourceType = GetType(NativeMethods.RASIPADDR) OrElse sourceType = GetType(String) Then
        '                Return True
        '#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
        '			ElseIf sourceType = GetType(NativeMethods.RASIPV6ADDR) Then
        '				Return True
        '#End If
        '#If (WIN7 OrElse WIN8) Then
        '			ElseIf sourceType = GetType(NativeMethods.RASTUNNELENDPOINT) Then
        '				Return True
        '			End If
        '#End If

        '                Return MyBase.CanConvertFrom(context, sourceType)
        '        End Function

        ''' <summary>
        ''' Returns whether this converter can convert an object to the destination type.
        ''' </summary>
        ''' <param name="context">An <see cref="ITypeDescriptorContext"/> that provides a format context.</param>
        ''' <param name="destinationType">The destination type.</param>
        ''' <returns><b>true</b> if the conversion can be performed, otherwise <b>false</b>.</returns>
        '        Public Overrides Function CanConvertTo(ByVal context As ITypeDescriptorContext, ByVal destinationType As Type) As Boolean
        '            If destinationType = GetType(NativeMethods.RASIPADDR) Then
        '                Return True
        '#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
        '			ElseIf destinationType = GetType(NativeMethods.RASIPV6ADDR) Then
        '				Return True
        '			End If
        '#End If

        '                Return MyBase.CanConvertTo(context, destinationType)
        '        End Function

        ''' <summary>
        ''' Converts the given object from the type of this converter. 
        ''' </summary>
        ''' <param name="context">An <see cref="ITypeDescriptorContext"/> that provides a format context.</param>
        ''' <param name="culture">A <see cref="System.Globalization.CultureInfo"/>. If null is passed, the current culture is presumed.</param>
        ''' <param name="value">The value to convert.</param>
        ''' <returns>The converted object.</returns>
        '        Public Overrides Function ConvertFrom(ByVal context As ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
        '            If TypeOf value Is String Then
        '                Return IPAddress.Parse(DirectCast(value, String))
        '            ElseIf TypeOf value Is NativeMethods.RASIPADDR Then
        '                Return New IPAddress(DirectCast(value, NativeMethods.RASIPADDR).addr)
        '#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
        '			ElseIf TypeOf value Is NativeMethods.RASIPV6ADDR Then
        '				Return New IPAddress(DirectCast(value, NativeMethods.RASIPV6ADDR).addr)
        '#End If
        '#If (WIN7 OrElse WIN8) Then
        '			ElseIf TypeOf value Is NativeMethods.RASTUNNELENDPOINT Then
        '				Dim endpoint As NativeMethods.RASTUNNELENDPOINT = DirectCast(value, NativeMethods.RASTUNNELENDPOINT)
        '				If endpoint.type <> NativeMethods.RASTUNNELENDPOINTTYPE.Unknown Then
        '					Select Case endpoint.type
        '						Case NativeMethods.RASTUNNELENDPOINTTYPE.IPv4
        '							Dim addr As Byte() = New Byte(3) {}
        '							Array.Copy(endpoint.addr, 0, addr, 0, 4)

        '							Return New IPAddress(addr)

        '						Case NativeMethods.RASTUNNELENDPOINTTYPE.IPv6
        '							Return New IPAddress(endpoint.addr)
        '					End Select
        '				End If

        '				Return Nothing
        '			End If
        '#End If

        '                Return MyBase.ConvertFrom(context, culture, value)
        '        End Function

        ''' <summary>
        ''' Converts the given object to the type of this converter.
        ''' </summary>
        ''' <param name="context">An <see cref="ITypeDescriptorContext"/> that provides a format context.</param>
        ''' <param name="culture">A <see cref="System.Globalization.CultureInfo"/>. If null is passed, the current culture is presumed.</param>
        ''' <param name="value">The value to convert.</param>
        ''' <param name="destinationType">The destination type.</param>
        ''' <returns>The converted object.</returns>
        '        Public Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object
        '            Dim addr As IPAddress = DirectCast(value, IPAddress)

        '            If destinationType = GetType(NativeMethods.RASIPADDR) AndAlso (addr Is Nothing OrElse (addr IsNot Nothing AndAlso addr.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork)) Then
        '                Dim ipv4 As New NativeMethods.RASIPADDR()

        '                If value Is Nothing Then
        '                    ipv4.addr = IPAddress.Any.GetAddressBytes()
        '                Else
        '                    ipv4.addr = addr.GetAddressBytes()
        '                End If

        '                Return ipv4
        '#If (WIN2K8 OrElse WIN7 OrElse WIN8) Then
        '			ElseIf destinationType = GetType(NativeMethods.RASIPV6ADDR) AndAlso (addr Is Nothing OrElse (addr IsNot Nothing AndAlso addr.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6)) Then
        '				Dim ipv6 As New NativeMethods.RASIPV6ADDR()

        '				If addr Is Nothing Then
        '					ipv6.addr = IPAddress.IPv6Any.GetAddressBytes()
        '				Else
        '					ipv6.addr = addr.GetAddressBytes()
        '				End If

        '				Return ipv6
        '#End If

        '#If (WIN7 OrElse WIN8) Then
        '			ElseIf destinationType = GetType(NativeMethods.RASTUNNELENDPOINT) Then
        '				Dim endpoint As New NativeMethods.RASTUNNELENDPOINT()

        '				If addr IsNot Nothing Then
        '					Dim bytes As Byte() = New Byte(15) {}
        '					Dim actual As Byte() = addr.GetAddressBytes()

        '					' Transfer the bytes to the 
        '					Array.Copy(actual, bytes, actual.Length)

        '					Select Case addr.AddressFamily
        '						Case AddressFamily.InterNetwork
        '							endpoint.type = NativeMethods.RASTUNNELENDPOINTTYPE.IPv4
        '							Exit Select

        '						Case AddressFamily.InterNetworkV6
        '							endpoint.type = NativeMethods.RASTUNNELENDPOINTTYPE.IPv6
        '							Exit Select
        '						Case Else

        '							endpoint.type = NativeMethods.RASTUNNELENDPOINTTYPE.Unknown
        '							Exit Select
        '					End Select

        '					endpoint.addr = bytes
        '				End If

        '				Return endpoint
        '			End If

        '#End If

        '                Return MyBase.ConvertTo(context, culture, value, destinationType)
        '        End Function

#End Region
    End Class
End Namespace
