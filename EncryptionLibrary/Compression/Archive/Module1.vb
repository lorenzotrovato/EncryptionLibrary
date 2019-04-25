Imports System.IO
Imports EncryptionLibrary.Compression.Archive
Module Module1

    Dim WithEvents cmp As ArchiveCompressor
    Dim WithEvents dcm As ArchiveDecompressor

    Dim fileout As New FileInfo(My.Computer.FileSystem.CombinePath(My.Computer.FileSystem.SpecialDirectories.Desktop, "proveCompression.zip"))
    Sub Main()
        Console.WindowWidth = Console.LargestWindowWidth
        's = New SymmetricCrypto("ciao1234", inputfile, outputfile)
        'a = New AsymmetricCrypto("provastringa", AsymmetricCrypto.AsymmetricAlgos.RSA_Key4096bits)
        'Console.WriteLine("generating keys ...")
        'Console.ReadKey()+

        If fileout.Exists Then
            fileout.Delete()
        End If

        Dim prova As New List(Of FileInfo)
        Dim dircontent = My.Computer.FileSystem.GetFiles("C:\Users\loren\Documents\banca")
        For Each f In dircontent
            prova.Add(New FileInfo(f))
        Next
        Console.WriteLine("Starting compression...")
        cmp = New ArchiveCompressor(prova, fileout)
        cmp.Start()
        Threading.Thread.Sleep(3500)
        'cmp.Cancel()
        Console.ReadKey()
    End Sub

#Region "CompressionArchive"
    Private Sub cmp_Progress(currentFileIndex As Integer, currentFileCompleted As Long, currentFileTotal As Long, nowWritten As Integer, totalFilesCompleted As Long, totalFilesTotal As Long) Handles cmp.Progress
        Dim percentSingle As Double = Math.Round((currentFileCompleted * 100) / currentFileTotal, 2)
        Dim percentTotal As Double = Math.Round((totalFilesCompleted * 100) / totalFilesTotal, 2)
        Console.Write("File N° " & (currentFileIndex + 1) & ": " & percentSingle & "% - Total: " & percentTotal & "% -- ")
        Console.Write("CurrentFileCompleted: " & currentFileCompleted & " - CurrentFileTotal: " & currentFileTotal & " -- ")
        Console.Write("TotalFileCompleted: " & totalFilesCompleted & " - TotalFileTotal: " & totalFilesTotal & " -- ")
        Console.WriteLine("Bytes written now: " & nowWritten)
    End Sub

    Private Sub cmp_Finished(totalCompleted As ULong) Handles cmp.Finished
        Console.WriteLine("Completed compression: " & totalCompleted & " bytes written")
        Console.ReadKey()
        Dim dirout As New DirectoryInfo("C:\Users\loren\Desktop\proveDecompress")
        For Each file In dirout.GetFiles()
            file.Delete()
        Next
        dcm = New ArchiveDecompressor(fileout, dirout)
        dcm.Start()
        Threading.Thread.Sleep(3500)
        dcm.Cancel()
    End Sub

    Private Sub cmp_Aborted() Handles cmp.Aborted
        Console.WriteLine("Compression aborted ---------------------")
        Console.ReadKey()
    End Sub

#End Region


#Region "DecompressionArchive"
    Private Sub dcm_Progress(currentFileIndex As Integer, currentFileCompleted As Long, currentFileTotal As Long, nowWritten As Integer, totalFilesCompleted As Long, totalFilesTotal As Long) Handles dcm.Progress
        Dim percentSingle As Double = Math.Round((currentFileCompleted * 100) / currentFileTotal, 2)
        Dim percentTotal As Double = Math.Round((totalFilesCompleted * 100) / totalFilesTotal, 2)
        Console.Write("File N° " & (currentFileIndex + 1) & ": " & percentSingle & "% - Total: " & percentTotal & "% -- ")
        Console.Write("CurrentFileCompleted: " & currentFileCompleted & " - CurrentFileTotal: " & currentFileTotal & " -- ")
        Console.Write("TotalFileCompleted: " & totalFilesCompleted & " - TotalFileTotal: " & totalFilesTotal & " -- ")
        Console.WriteLine("Bytes written now: " & nowWritten)
    End Sub

    Private Sub dcm_Finished(totalCompleted As ULong) Handles dcm.Finished
        Console.WriteLine("Completed decompression: " & totalCompleted & " bytes written")
        Console.ReadKey()
    End Sub

    Private Sub dcm_Aborted() Handles dcm.Aborted
        Console.WriteLine("Decompression aborted ---------------------")
        Console.ReadKey()
    End Sub
#End Region

End Module
