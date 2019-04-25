Imports System.IO
Imports System.IO.Compression
Imports System.Threading


Namespace Compression

    ''' <summary>
    ''' Provide methods for decompressing a file from a ZIP archive
    ''' </summary>
    Public Class Decompressor
        Implements IDisposable

        'events
        ''' <summary>
        '''     <para>Fired when a block of the stream has been decompressed</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="bytesDecompressed">Total bytes decompressed until now</param>
        ''' <param name="bytesTotal">Total bytes to write</param>
        ''' <param name="nowWritten">Bytes written in this step</param>
        Public Event Progress(ByVal bytesDecompressed As ULong, ByVal bytesTotal As ULong, ByVal nowWritten As UInteger)

        ''' <summary>
        '''     <para>Fired when the decompression is completed</para>
        '''     <para>The object is automatically disposed (cannot be used anymore)</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="bytesWritten">Total bytes decompressed</param>
        Public Event Finished(ByVal bytesWritten As ULong)

        ''' <summary>
        '''     <para>Fired when the process has been stopped or the file does not exists</para>
        '''     <para>The object is automatically disposed (cannot be used anymore)</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        Public Event Aborted()


        'I/O files
        Private inputFileStream As FileStream
        Private fileToExtract As String
        Private outputFileStream As FileStream
        Private archive As ZipArchive
        'thread
        Private decompressionThread As Thread
        'written length
        Private completedLength As ULong = 0
        'abort operation
        Private abort As Boolean = False


        ''' <summary>
        ''' Initializes a new file decompressor
        ''' </summary>
        ''' <param name="input">ZIP archive file</param>
        ''' <param name="filenameToExtract">Name of the file to be extracted</param>
        ''' <param name="outputPath">Decompressed file path</param>
        Public Sub New(ByVal input As FileInfo, ByVal filenameToExtract As String, ByVal outputPath As FileInfo)
            Me.inputFileStream = New FileStream(input.FullName, FileMode.Open, FileAccess.Read)
            Me.archive = New ZipArchive(Me.inputFileStream, ZipArchiveMode.Read)
            Me.outputFileStream = New FileStream(outputPath.FullName, FileMode.Create, FileAccess.Write)
            Me.decompressionThread = New Thread(New ThreadStart(AddressOf Decompress))
            Me.completedLength = 0
            Me.fileToExtract = filenameToExtract
        End Sub

        ''' <summary>
        ''' Starts the asynchronous decompression
        ''' </summary>
        Public Sub StartDecompression()
            If (Not Me.abort) Then
                Me.decompressionThread.Start()
            End If
        End Sub

        ''' <summary>
        ''' Stops the decompression and automatically dispose the object
        ''' </summary>
        Public Sub Cancel()
            Me.abort = True
        End Sub

        ''' <summary>
        ''' Flush, close and release all resources used by the streams
        ''' </summary>
        ''' See <see cref="IDisposable.Dispose()"/>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Me.inputFileStream.Flush()
                Me.archive.Dispose()
                Me.outputFileStream.Flush()
                Me.inputFileStream.Close()
                Me.outputFileStream.Close()
                Me.inputFileStream.Dispose()
                Me.outputFileStream.Dispose()
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' Gets the original file passed to the constructor
        ''' </summary>
        ''' <returns>The original FileInfo instance</returns>
        Public ReadOnly Property InputFile As FileInfo
            Get
                Return New FileInfo(Me.inputFileStream.Name)
            End Get
        End Property

        'async decompression
        Private Sub Decompress()
            Dim buffer() As Byte
            Dim currentRead As UInteger = 0
            Dim archiveEntry As ZipArchiveEntry = Nothing
            For Each en In Me.archive.Entries
                If (en.Name = Me.fileToExtract) Then
                    archiveEntry = Me.archive.GetEntry(en.Name)
                End If
            Next
            If (archiveEntry IsNot Nothing) Then
                Dim entryStream As Stream = archiveEntry.Open()
                Debug.WriteLine(entryStream.Length)
                buffer = New Byte(Encryption.Utils.CalculateCopyBufferLength(Me.inputFileStream.Length)) {}
                While (Me.completedLength < archiveEntry.Length And Not Me.abort)
                    currentRead = entryStream.Read(buffer, 0, buffer.Length)
                    Me.outputFileStream.Write(buffer, 0, currentRead)
                    Me.completedLength += currentRead
                    RaiseEvent Progress(Me.completedLength, archiveEntry.Length, currentRead)
                End While
                entryStream.Flush()
                entryStream.Close()
                Me.Dispose()
                If (Me.abort) Then
                    RaiseEvent Aborted()
                Else
                    RaiseEvent Finished(Me.completedLength)
                End If
            Else
                Me.Dispose()
                RaiseEvent Aborted()
            End If
        End Sub
    End Class
End Namespace
