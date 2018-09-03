
Imports System.Runtime.InteropServices

Public Class ClassDriverEvent


    Public Declare Auto Function SetDifxLogCallback Lib "DIFxAPI.dll" (ByVal LogCallback As DIFLOGCALLBACK, ByVal CallbackContext As IntPtr) As Int32

    Public Delegate Sub DIFLOGCALLBACK(ByVal EventType As DIFXAPI_LOG, ByVal ErrorCode As Int32, <MarshalAs(UnmanagedType.LPTStr)> ByVal EventDescription As String, ByVal CallbackContext As IntPtr)
    Public Declare Auto Function DIFXAPISetLogCallback Lib "DIFxAPI.dll" (ByVal LogCallback As DIFXAPILOGCALLBACK, ByVal CallbackContext As IntPtr) As Int32


    Public Delegate Sub DIFXAPILOGCALLBACK(ByVal EventType As DIFXAPI_LOG, ByVal ErrorCode As Int32, <MarshalAs(UnmanagedType.LPTStr)> ByVal EventDescription As String, ByVal CallbackContext As IntPtr)

    Public Enum DIFXAPI_LOG
        DIFXAPI_SUCCESS = 0  ' Successes
        DIFXAPI_INFO = 1     ' Basic logging information that should always be shown
        DIFXAPI_WARNING = 2  ' Warnings
        DIFXAPI_ERROR = 3    ' Errors
    End Enum

    Public Shared Sub Register()
        Try

            '   SetDifxLogCallback(New DIFLOGCALLBACK(AddressOf DIFLogCallbackFunc), IntPtr.Zero)

            dIFXAPISetLogCallback(New DIFXAPILOGCALLBACK(AddressOf DIFxAPILogCallbackFunc), IntPtr.Zero)

        Catch ex As Exception
            '   MessageBox.Show(ex.Message, "ClassDriverEvent.Register", MessageBoxButtons.OK)
        End Try
    End Sub

    'Public Shared Sub DIFxAPILogCallbackFunc(ByVal EventType As DIFXAPI_LOG, ByVal ErrorCode As Int32, <MarshalAs(UnmanagedType.LPTStr)> ByVal EventDescription As String, ByVal CallbackContext As IntPtr)
    '    Try
    '        MessageBox.Show(EventDescription, EventType.ToString)
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "ClassDriverEvent.DIFLogCallbackFunc", MessageBoxButtons.OK)
    '    End Try
    'End Sub



    Private Shared Sub DIFLogCallbackFunc(ByVal EventType As DIFXAPI_LOG, ByVal ErrorCode As Int32, <MarshalAs(UnmanagedType.LPTStr)> ByVal EventDescription As String, ByVal CallbackContext As IntPtr)

        Try
            MessageBox.Show(EventDescription, EventType.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ClassDriverEvent.DIFLogCallbackFunc", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Shared Sub DIFxAPILogCallbackFunc(ByVal EventType As DIFXAPI_LOG, ByVal ErrorCode As Int32, ByVal EventDescription As String, ByVal CallbackContext As IntPtr)

        Try
            Dim s As String = ""
       

            Select Case EventType
                Case DIFXAPI_LOG.DIFXAPI_SUCCESS
                    s = String.Format("SUCCESS: {0}. Error code: {1}", EventDescription, ErrorCode)
                    Exit Sub
                Case DIFXAPI_LOG.DIFXAPI_INFO
                    s = String.Format("INFO: {0}. Error code: {1}", EventDescription, ErrorCode)
                    Exit Sub
                Case DIFXAPI_LOG.DIFXAPI_WARNING
                    s = String.Format("WARNING: {0}. Error code: {1}", EventDescription, ErrorCode)
                    Exit Sub
                Case DIFXAPI_LOG.DIFXAPI_ERROR
                    s = String.Format("ERROR: {0}. Error code: {1}", EventDescription, ErrorCode)
                    Exit Sub
            End Select

        Catch ex As Exception

        End Try
    End Sub

End Class
