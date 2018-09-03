Imports Microsoft.Win32.SafeHandles


Friend Class SafeDeviceInfoSetHandle
    Inherits SafeHandleZeroOrMinusOneIsInvalid

    Sub New()
        MyBase.New(True)
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        NativeMethods1.SetupDiDestroyDeviceInfoList(Me.handle)
    End Function

End Class
