Public Class ClassImpostazioniMcp
    Public Const NomeDirMcp As String = "Mcp"
    Public Const nomefile As String = "LastPathMcp.ini"

    Public Shared Sub StartCreaPath()
        Try
            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf _saveActualPath)
        Catch ex As Exception

        End Try
    End Sub

    Private Shared Sub _saveActualPath(ByVal o As Object)
        Try
            Dim p As String = LibPaths.ClassPATH.CreaDirAppdataMcp

            If Not My.Computer.FileSystem.DirectoryExists(p) Then
                My.Computer.FileSystem.CreateDirectory(p)
            End If

            Dim f As String = My.Computer.FileSystem.CombinePath(p, LibPaths.ClassPATH.LastPathMcp)

            Dim iniObj As New Ini.IniFile(f)

            Dim c As String = My.Application.Info.DirectoryPath

            iniObj.IniWriteValue("LastApplicationPath", LibPaths.ClassPATH.NomeCartellaProgramma, c)

        Catch ex As Exception

        End Try

    End Sub

    Public Shared Function LoadActualPath(ByRef path As String) As Boolean
        Try
            path = ""

            Dim p As String = LibPaths.ClassPATH.CreaDirAppdataMcp

            Dim f As String = My.Computer.FileSystem.CombinePath(p, LibPaths.ClassPATH.LastPathMcp)

            Dim iniObj As New Ini.IniFile(f)

            Dim c As String = My.Application.Info.DirectoryPath

            Dim val As String = iniObj.IniReadValue("LastApplicationPath", LibPaths.ClassPATH.NomeCartellaProgramma)

            If val IsNot Nothing AndAlso val.Trim <> "" Then
                path = val.Trim
                Return True
            End If

        Catch ex As Exception

        End Try

    End Function

End Class
