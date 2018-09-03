#Region "References"

Imports System.Runtime.InteropServices

#End Region

Public Class LaunchProcess

#Region "Native Methods"

    <DllImport("User32.dll", SetLastError:=True)> Private Shared Function GetShellWindow() As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> Private Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)> Private Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As IntPtr) As Integer
    End Function

    <DllImport("kernel32.dll")> Private Shared Function OpenProcess(ByVal dwDesiredAccess As UInteger, <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, ByVal dwProcessId As IntPtr) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function DuplicateTokenEx(
    ByVal ExistingTokenHandle As IntPtr,
    ByVal dwDesiredAccess As UInt32,
    ByRef lpThreadAttributes As SECURITY_ATTRIBUTES,
    ByVal ImpersonationLevel As Integer,
    ByVal TokenType As Integer,
    ByRef DuplicateTokenHandle As System.IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> Private Shared Function LookupPrivilegeValue(lpSystemName As String, lpName As String, ByRef lpLuid As LUID) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function AdjustTokenPrivileges(
    ByVal TokenHandle As IntPtr,
    ByVal DisableAllPrivileges As Boolean,
    ByRef NewState As TOKEN_PRIVILEGES,
    ByVal BufferLengthInBytes As Integer,
    ByRef PreviousState As TOKEN_PRIVILEGES,
    ByRef ReturnLengthInBytes As Integer
  ) As Boolean
    End Function

    <DllImport("advapi32", SetLastError:=True, CharSet:=CharSet.Unicode)> Private Shared Function CreateProcessWithTokenW(hToken As IntPtr, dwLogonFlags As Integer, lpApplicationName As String, lpCommandLine As String, dwCreationFlags As Integer, lpEnvironment As IntPtr, lpCurrentDirectory As IntPtr, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
    End Function

#End Region

#Region "Structures"

    <StructLayout(LayoutKind.Sequential)> Private Structure SECURITY_ATTRIBUTES

        Friend nLength As Integer
        Friend lpSecurityDescriptor As IntPtr
        Friend bInheritHandle As Integer

    End Structure

    Private Structure TOKEN_PRIVILEGES

        Friend PrivilegeCount As Integer
        Friend TheLuid As LUID
        Friend Attributes As Integer

    End Structure

    Private Structure LUID

        Friend LowPart As UInt32
        Friend HighPart As UInt32

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> Private Structure STARTUPINFO

        Friend cb As Integer
        Friend lpReserved As String
        Friend lpDesktop As String
        Friend lpTitle As String
        Friend dwX As Integer
        Friend dwY As Integer
        Friend dwXSize As Integer
        Friend dwYSize As Integer
        Friend dwXCountChars As Integer
        Friend dwYCountChars As Integer
        Friend dwFillAttribute As Integer
        Friend dwFlags As Integer
        Friend wShowWindow As Short
        Friend cbReserved2 As Short
        Friend lpReserved2 As Integer
        Friend hStdInput As Integer
        Friend hStdOutput As Integer
        Friend hStdError As Integer

    End Structure

    Private Structure PROCESS_INFORMATION

        Friend hProcess As IntPtr
        Friend hThread As IntPtr
        Friend dwProcessId As Integer
        Friend dwThreadId As Integer

    End Structure

#End Region

#Region "Enumerations"

    Private Enum TOKEN_INFORMATION_CLASS

        TokenUser = 1
        TokenGroups
        TokenPrivileges
        TokenOwner
        TokenPrimaryGroup
        TokenDefaultDacl
        TokenSource
        TokenType
        TokenImpersonationLevel
        TokenStatistics
        TokenRestrictedSids
        TokenSessionId
        TokenGroupsAndPrivileges
        TokenSessionReference
        TokenSandBoxInert
        TokenAuditPolicy
        TokenOrigin
        TokenElevationType
        TokenLinkedToken
        TokenElevation
        TokenHasRestrictions
        TokenAccessInformation
        TokenVirtualizationAllowed
        TokenVirtualizationEnabled
        TokenIntegrityLevel
        TokenUIAccess
        TokenMandatoryPolicy
        TokenLogonSid
        MaxTokenInfoClass

    End Enum

#End Region

#Region "Constants"

    Private Const SE_PRIVILEGE_ENABLED As Integer = &H2L
    Private Const PROCESS_QUERY_INFORMATION As Integer = &H400
    Private Const TOKEN_ASSIGN_PRIMARY As Integer = &H1
    Private Const TOKEN_DUPLICATE As Integer = &H2
    Private Const TOKEN_IMPERSONATE As Integer = &H4
    Private Const TOKEN_QUERY As Integer = &H8
    Private Const TOKEN_QUERY_SOURCE As Integer = &H10
    Private Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
    Private Const TOKEN_ADJUST_GROUPS As Integer = &H40
    Private Const TOKEN_ADJUST_DEFAULT As Integer = &H80
    Private Const TOKEN_ADJUST_SESSIONID As Integer = &H100
    Private Const SecurityImpersonation As Integer = 2
    Private Const TokenPrimary As Integer = 1
    Private Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"

#End Region

#Region "Methods"

    Friend Shared Function LaunchFile(ByVal FilePath As String, ByVal IsWaitToFinish As Boolean) As Boolean

        Try

            'Enable the SeIncreaseQuotaPrivilege in current token
            Dim HPrcsToken As IntPtr = Nothing
            OpenProcessToken(Process.GetCurrentProcess.Handle, TOKEN_ADJUST_PRIVILEGES, HPrcsToken)

            Dim TokenPrvlgs As TOKEN_PRIVILEGES
            TokenPrvlgs.PrivilegeCount = 1
            LookupPrivilegeValue(Nothing, SE_INCREASE_QUOTA_NAME, TokenPrvlgs.TheLuid)
            TokenPrvlgs.Attributes = SE_PRIVILEGE_ENABLED

            AdjustTokenPrivileges(HPrcsToken, False, TokenPrvlgs, 0, Nothing, Nothing)

            'Get window handle representing the desktop shell
            Dim HShellWnd As IntPtr = GetShellWindow()

            'Get the ID of the desktop shell process
            Dim ShellPID As IntPtr
            GetWindowThreadProcessId(HShellWnd, ShellPID)

            'Open the desktop shell process in order to get the process token
            Dim HShellPrcs As IntPtr = OpenProcess(PROCESS_QUERY_INFORMATION, False, ShellPID)
            Dim HShellPrcSToken As IntPtr = Nothing
            Dim HPrimaryToken As IntPtr = Nothing

            'Get the process token of the desktop shell
            OpenProcessToken(HShellPrcs, TOKEN_DUPLICATE, HShellPrcSToken)

            'Duplicate the shell's process token to get a primary token
            Dim TokenRights As UInteger = TOKEN_QUERY Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID
            DuplicateTokenEx(HShellPrcSToken, TokenRights, Nothing, SecurityImpersonation, TokenPrimary, HPrimaryToken)

            Dim StartInfo As STARTUPINFO = Nothing
            Dim PrcsInfo As PROCESS_INFORMATION = Nothing

            StartInfo.cb = Marshal.SizeOf(StartInfo)
            Dim IsSuccessed As Boolean = CreateProcessWithTokenW(HPrimaryToken, 1, FilePath, "", 0, Nothing, Nothing, StartInfo, PrcsInfo)

            If IsSuccessed = True Then

                If IsWaitToFinish = True Then

                    Try

                        Dim Prcs As Process = Process.GetProcessById(PrcsInfo.dwProcessId)
                        Prcs.WaitForExit()

                    Catch ex As Exception
                    End Try

                End If

            Else

                'Could not launch process with shell token may be the process needs admin rights to launch, we will try to launch it with default parent process permissions.

                Dim Prcs As New Process
                Prcs.StartInfo.FileName = FilePath
                Prcs.Start()

                If IsWaitToFinish = True Then Prcs.WaitForExit()

            End If

            Return True

        Catch ex As Exception
        End Try

    End Function

#End Region

End Class