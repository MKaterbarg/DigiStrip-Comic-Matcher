Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Collections.Generic
Imports System.Text.Json
Imports RestSharp
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Program
    Sub Main(args As String())

        Dim MovedCounter = 0
        Dim FailedCounter = 0
        Dim RemovedCounter = 0

        If args.Count < 2 Then

            Console.WriteLine("Please provide both the source path and destination path as arguments.")
            Console.WriteLine("Usage: DigiStripComicMatcher.exe SourcePath DestinationPath")
            Console.WriteLine("All files in the Source Path, and subdirectories, will be scanned and based on their CRC32 hash be moved to the Destination Path.")
            Console.WriteLine("")
            Console.WriteLine("For more detailed information, please use: DigiStripComicMatcher.exe -help")
            Exit Sub

        End If
        Dim RemoveDuplicate As Boolean = False
        For Each arg In args
            If arg = "-Help" Then

                Help()
                Exit Sub

            End If

            If arg = "-RemoveDuplicate" Then

                RemoveDuplicate = True

            End If
        Next

        Dim sourcePath = args(0)
        Dim destPath = args(1)

        If Not IO.Directory.Exists(sourcePath) Then
            Console.WriteLine("The source directory " + sourcePath + " does not exist.")
            Exit Sub
        End If

        If Not IO.Directory.Exists(destPath) Then
            Console.WriteLine("The destination directory " + destPath + " does not exist.")
            Exit Sub
        End If

        'Scan sourcePath for new *.cbr files
        Console.WriteLine("Please wait while we index the files found in the source path. This may take a few minutes depending on the number of files.")
        Dim cbrFilesArray As String() = Directory.GetFiles(sourcePath, "*.cbr", SearchOption.AllDirectories)
        Console.WriteLine("Total Number of files found: " + cbrFilesArray.Count.ToString)

        For Each file In cbrFilesArray

            Dim CRC = GetCRC32(file)
            Dim f = New IO.FileInfo(file)
            Console.WriteLine("File Found: " + f.Name)
            Console.WriteLine("CRC:" + CRC)

            If CRC.Length > 0 Then

                Dim FoundIssue As Linq.JArray = FindByCRC32(CRC)

                If FoundIssue.Count = 1 Then

                    Console.WriteLine("Match Found. Moving file")

                    If Not Directory.Exists(destPath + "/" + FoundIssue(0)("comic")("name").ToString) Then
                        Try
                            MkDir(destPath + "/" + FoundIssue(0)("comic")("name").ToString)
                        Catch ex As Exception

                            Console.WriteLine("Unable to create directory " + destPath + "/" + FoundIssue(0)("comic")("name").ToString)
                            Exit Sub
                        End Try
                    End If

                    If Not IO.File.Exists(destPath + "/" + FoundIssue(0)("comic")("name").ToString + "/" + FoundIssue(0)("filename").ToString) Then
                        Try
                            f.MoveTo(destPath + "/" + FoundIssue(0)("comic")("name").ToString + "/" + FoundIssue(0)("filename").ToString)
                            MovedCounter += 1
                        Catch ex As Exception

                            FailedCounter += 1
                        End Try
                    Else
                        If RemoveDuplicate = True Then
                            f.Delete()
                            RemovedCounter += 1
                        Else
                            FailedCounter += 1
                        End If

                    End If

                Else
                    FailedCounter += 1
                End If
            Else

                FailedCounter += 1
            End If


        Next

        Console.WriteLine("All Done. " + vbNewLine +
                          "Failed: " + FailedCounter.ToString + vbNewLine +
                          "Succes: " + MovedCounter.ToString + vbNewLine +
                          "Removed: " + RemovedCounter.ToString)


    End Sub

    Private Sub Help()
        Throw New NotImplementedException()
    End Sub

    Public Function FindByCRC32(CRC32 As String) As JArray

        Dim client = New RestClient("https://nl-comics-api.jag.digital")

        Dim request = New RestRequest("/findByCRC32/" + CRC32, DataFormat.Json)

        Dim response = client.Get(request)

        Return JsonConvert.DeserializeObject(response.Content)

    End Function
    Public Function GetCRC32(ByVal sFileName As String) As String

        Try
            Dim FS As FileStream = New FileStream(sFileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            Dim CRC32Result As Integer = &HFFFFFFFF
            Dim Buffer(4096) As Byte
            Dim ReadSize As Integer = 4096
            Dim Count As Integer = FS.Read(Buffer, 0, ReadSize)
            Dim CRC32Table(256) As Integer
            Dim DWPolynomial As Integer = &HEDB88320
            Dim DWCRC As Integer
            Dim i As Integer, j As Integer, n As Integer
            Dim Dinges As String = ""
            Dim strCRC32 As String = ""
            Dim blnCRC32 As Boolean = False

            'Create CRC32 Table
            For i = 0 To 255
                DWCRC = i
                For j = 8 To 1 Step -1
                    If (DWCRC And 1) Then
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                        DWCRC = DWCRC Xor DWPolynomial
                    Else
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    End If
                Next j
                CRC32Table(i) = DWCRC
            Next i

            'Calcualting CRC32 Hash
            Do While (Count > 0)
                For i = 0 To Count - 1
                    n = (CRC32Result And &HFF) Xor Buffer(i)
                    CRC32Result = ((CRC32Result And &HFFFFFF00) \ &H100) And &HFFFFFF
                    CRC32Result = CRC32Result Xor CRC32Table(n)
                Next i
                Count = FS.Read(Buffer, 0, ReadSize)
            Loop

            Dinges = (Hex(Not (CRC32Result)))
            FS.Close()

            Select Case Dinges.Length

                Case 4
                    Return "0000" + Dinges.ToString


                Case 5
                    Return "000" + Dinges.ToString


                Case 6
                    Return "00" + Dinges.ToString


                Case 7
                    Return "0" + Dinges.ToString


                Case 8
                    Return Dinges.ToString


                Case Else
                    Return ""

            End Select

        Catch ex As Exception
            Return ""

        End Try

    End Function

End Module
