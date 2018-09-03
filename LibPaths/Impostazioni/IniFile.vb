Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Ini
    ''' <summary>
    ''' Create a New INI file to store or load data
    ''' </summary>
    Public Class IniFile
        Public path As String

        <DllImport("kernel32")> _
        Private Shared Function WritePrivateProfileString(ByVal section As String, ByVal key As String, ByVal val As String, ByVal filePath As String) As Long
        End Function
        <DllImport("kernel32")> _
        Private Shared Function GetPrivateProfileString(ByVal section As String, ByVal key As String, ByVal def As String, ByVal retVal As StringBuilder, ByVal size As Integer, ByVal filePath As String) As Integer
        End Function

        ''' <summary>
        ''' INIFile Constructor.
        ''' </summary>
        ''' <param name="INIPath"></param>
        Public Sub New(ByVal INIPath As String)
            path = INIPath
        End Sub
        ''' <summary>
        ''' Write Data to the INI File
        ''' </summary>
        ''' <param name="Section"></param>
        ''' Section name
        ''' <param name="Key"></param>
        ''' Key Name
        ''' <param name="Value"></param>
        ''' Value Name
        Public Sub IniWriteValue(ByVal Section As String, ByVal Key As String, ByVal Value As String)
            WritePrivateProfileString(Section, Key, Value, Me.path)
        End Sub

        ''' <summary>
        ''' Read Data Value From the Ini File
        ''' </summary>
        ''' <param name="Section"></param>
        ''' <param name="Key"></param>
        ''' <param name="Path"></param>
        ''' <returns></returns>
        Public Function IniReadValue(ByVal Section As String, ByVal Key As String) As String
            Dim temp As New StringBuilder(255)
            Dim i As Integer = GetPrivateProfileString(Section, Key, "", temp, 255, Me.path)
            Return temp.ToString()

        End Function
    End Class
End Namespace