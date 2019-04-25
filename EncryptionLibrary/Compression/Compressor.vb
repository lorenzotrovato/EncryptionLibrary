Imports System.IO
Imports System.IO.Compression
Imports System.Threading


Namespace Compression

    ''' <summary>
    ''' Provide methods for compressing a file into a ZIP archive
    ''' </summary>
    Public Class Compressor
        Implements IDisposable

        'events
        ''' <summary>
        '''     <para>Fired when a block of the stream has been compressed</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="bytesCompressed">Total bytes compressed until now</param>
        ''' <param name="bytesTotal">Total bytes to write</param>
        ''' <param name="nowWritten">Bytes written in this step</param>
        Public Event Progress(ByVal bytesCompressed As ULong, ByVal bytesTotal As ULong, ByVal nowWritten As UInteger)

        ''' <summary>
        '''     <para>Fired when the compression is completed</para>
        '''     <para>The object is automatically disposed (cannot be used anymore)</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="bytesWritten">Total bytes written</param>
        Public Event Finished(ByVal bytesWritten As ULong)

        ''' <summary>
        '''     <para>Fired when the process has been stopped</para>
        '''     <para>The object is automatically disposed (cannot be used anymore)</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        Public Event Aborted()


        'I/O files
        Private inputFileStream As FileStream
        Private outputFileStream As FileStream
        Private archive As ZipArchive
        'thread
        Private compressionThread As Thread
        'written length
        Private completedLength As ULong = 0
        'abort operation
        Private abort As Boolean = False

        ''' <summary>
        ''' Initializes a new file compressor
        ''' </summary>
        ''' <param name="input">File to be compressed</param>
        ''' <param name="outputPath">Path of the archive</param>
        ''' <param name="overwrite">Tells if the file, if present, has to be overwritten (default FALSE)</param>
        Public Sub New(ByVal input As FileInfo, ByVal outputPath As FileInfo, ByVal Optional overwrite As Boolean = False)
            Me.inputFileStream = New FileStream(input.FullName, FileMode.Open, FileAccess.Read)
            If (Not overwrite And outputPath.Exists) Then
                Me.outputFileStream = New FileStream(outputPath.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                Me.archive = New ZipArchive(Me.outputFileStream, ZipArchiveMode.Update)
            Else
                Me.outputFileStream = New FileStream(outputPath.FullName, FileMode.Create, FileAccess.Write)
                Me.archive = New ZipArchive(Me.outputFileStream, ZipArchiveMode.Create)
            End If
            Me.compressionThread = New Thread(New ThreadStart(AddressOf Compress))
            Me.completedLength = 0
        End Sub

        ''' <summary>
        ''' Starts the asynchronous compression
        ''' </summary>
        Public Sub StartCompression()
            If (Not Me.abort) Then
                Me.compressionThread.Start()
            End If
        End Sub

        ''' <summary>
        ''' Stops the compression and automatically dispose the object
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
                Me.archive.Dispose()
                Me.outputFileStream.Flush()
                Me.outputFileStream.Close()
                Me.inputFileStream.Flush()
                Me.inputFileStream.Close()
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

        'async compression
        Private Sub Compress()
            Dim buffer(Encryption.Utils.CalculateCopyBufferLength(Me.inputFileStream.Length)) As Byte
            Dim currentRead As UInteger = 0
            Dim nEntry As ZipArchiveEntry = Me.archive.CreateEntry(New FileInfo(Me.inputFileStream.Name).Name, CompressionLevel.Fastest)
            Dim entryStream As Stream = nEntry.Open()
            While (Me.completedLength < Me.inputFileStream.Length And Not Me.abort)
                currentRead = Me.inputFileStream.Read(buffer, 0, buffer.Length)
                entryStream.Write(buffer, 0, currentRead)
                Me.completedLength += currentRead
                RaiseEvent Progress(Me.completedLength, Me.inputFileStream.Length, currentRead)
            End While
            entryStream.Flush()
            entryStream.Close()
            Me.Dispose()
            If (Me.abort) Then
                RaiseEvent Aborted()
            Else
                RaiseEvent Finished(Me.completedLength)
            End If
        End Sub
    End Class
End Namespace
