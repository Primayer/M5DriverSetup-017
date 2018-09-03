
Imports System.IO

Public Class ClassIni

    Private _path As String = ""

    Public Function GetValue(ByVal Gruppo As String, ByVal key As String, ByRef value As String) As Boolean
        Try
            Dim r As String = ""
            Dim groupfound As Boolean = False

            Using sr As New StreamReader(Me._path)

                Do While Not sr.EndOfStream

                    r = sr.ReadLine()

                    If r.IndexOf("["c) >= 0 AndAlso r.IndexOf("]"c) >= 0 AndAlso r.IndexOf(Gruppo) >= 0 Then
                        groupfound = True
                    End If

                    If groupfound AndAlso r.IndexOf("["c) < 0 AndAlso r.IndexOf("]"c) < 0 AndAlso r.IndexOf(key) >= 0 AndAlso r.IndexOf("="c) >= 0 Then
                        value = r.Split("="c)(1).Trim
                        Return True
                    End If

                Loop

            End Using


            Return False
        Catch ex As Exception

        End Try


    End Function


    Public Function GetEntries() As List(Of String)

        Try
            Dim r As String = ""

            Dim l As New List(Of String)
            Dim en As String = ""

            Using sr As New StreamReader(Me._path)

                Do While Not sr.EndOfStream

                    r = sr.ReadLine()

                    If r.IndexOf("["c) >= 0 AndAlso r.IndexOf("]"c) >= 0 Then

                        en = r.Replace("["c, "")
                        en = en.Replace("]"c, "")

                        l.Add(en)
                    End If
                Loop

            End Using


            Return l
        Catch ex As Exception

        End Try


    End Function

    Public Sub New(ByVal path As String)
        Me._path = path
    End Sub
End Class
