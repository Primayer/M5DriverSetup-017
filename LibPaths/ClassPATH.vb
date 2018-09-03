Public Class ClassPATH

#If MCPVER = "PRO" Then
    Public Const NomeCartellaProgramma As String = "McpPro"
    Public Const NomeProcessoMcp As String = "McpPro"
    Public Const LocalFTPPathInstall As String = "\\HEMINASVR01\Tecnico\BACKUP\MOSE'\Projects\M5\UpdateMcp\LocalServerFtpMcpPro\"
    Public Const CartellaFTP_Mcp As String = "McpPro"
    Public Const LastPathMcp As String = "LastPathMcpPro.ini"
    Public Const NomeFileInfoMcp As String = "InfoMcpPro.txt"
    
#Else
  Public Const NomeCartellaProgramma As String = "Mcp"
  Public Const NomeProcessoMcp As String = "Mcp"
  Public Const LocalFTPPathInstall As String = "\\HEMINASVR01\Tecnico\BACKUP\MOSE'\Projects\M5\UpdateMcp\LocalServerFtpMcp\"
  Public Const CartellaFTP_Mcp As String = "Mcp"
  Public Const LastPathMcp As String = "LastPathMcp.ini"
Public Const NomeFileInfoMcp As String = "InfoMcp.txt"


#End If

    Public Const NomeFileUpdater As String = "UpdateMcp.exe"


    Public Const NomeFile_LibDownloaddll As String = "LibDownloadMcp.dll"
    Public Const NomeFile_LibPathsdll As String = "LibPaths.dll"
    Public Const NomeFile_SPBdll As String = "SPB.dll"

    Public Const NomeFileLogUpdate As String = "LogUpdateMcp.txt"
    Public Const NomeFileLogDownload As String = "LogDownloadMcp.txt"
    Public Const NomeCartellaDownload As String = "Download"

    Public Const NomeFileManifestMcp As String = "ManifestMcp.ini"
    Public Const NomeFileManifestFtdi As String = "ManifestFtdi.ini"
  
   


    Public Const CartellaFTP_Log As String = "Log"
    Public Const CartellaFTP_Ftdi As String = "FTDI"


    Public Shared Function CreaDirAppdataMcp() As String
        Dim c As String = ""
        Try
            c = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\" + NomeCartellaProgramma + "\"

            If Not My.Computer.FileSystem.DirectoryExists(c) Then
                My.Computer.FileSystem.CreateDirectory(c)
            End If

            Return c

        Catch ex As Exception
            MessageBox.Show(c + vbCrLf + vbCrLf + ex.Message, "Create folder:", MessageBoxButtons.OK)
        End Try
    End Function

    Public Shared Function CreaDirAppdataDownload() As String
        Dim c As String = ""
        Try
            c = CreaDirAppdataMcp() + NomeCartellaDownload + "\"

            If Not My.Computer.FileSystem.DirectoryExists(c) Then
                My.Computer.FileSystem.CreateDirectory(c)
            End If

            Return c

        Catch ex As Exception
            MessageBox.Show(c + vbCrLf + vbCrLf + ex.Message, "Create folder:", MessageBoxButtons.OK)
        End Try

    End Function

    Public Shared Function GetPathM5USBList() As String
        Try
            Return My.Computer.FileSystem.CombinePath(CreaDirAppdataMcp, "M5Driver")
        Catch ex As Exception

        End Try
    End Function


    Public Shared Function GetPathM5Driver() As String
        Try
            Return My.Computer.FileSystem.CombinePath(GetPathDownload, "M5Driver")
        Catch ex As Exception

        End Try
    End Function

    Public Shared Function GetPathDownload() As String
        Try
            Return My.Computer.FileSystem.CombinePath(CreaDirAppdataMcp, NomeCartellaDownload)
        Catch ex As Exception

        End Try
    End Function



    Public Shared Function GetLastPathMcp() As String
        Try
            ' carica il file con l'ultimo percorso dell'applicazione
            Dim LoadActualPath As String = ""
            If ClassImpostazioniMcp.LoadActualPath(LoadActualPath) Then
                Return LoadActualPath
            Else
                Return CreaDirAppdataMcp()
            End If

        Catch ex As Exception

        End Try

    End Function

End Class
