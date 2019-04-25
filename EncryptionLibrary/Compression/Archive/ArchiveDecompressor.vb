Imports System.IO
Imports System.IO.Compression

Namespace Compression.Archive

    ''' <summary>
    ''' Extracts all files from a ZIP archive
    ''' </summary>
    Public Class ArchiveDecompressor
        Implements IDisposable

        ''' <summary>
        '''     <para>Fired when a block of one of the streams has been decompressed</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="currentFileIndex">Working file index based on the input list</param>
        ''' <param name="currentFileCompleted">Working file completed bytes</param>
        ''' <param name="currentFileTotal">Working file total size</param>
        ''' <param name="nowWritten">Bytes written in this step</param>
        ''' <param name="totalFilesCompleted">Total completed bytes</param>
        ''' <param name="totalFilesTotal">Total file sizes</param>
        Public Event Progress(ByVal currentFileIndex As Integer, ByVal currentFileCompleted As Long, ByVal currentFileTotal As Long, ByVal nowWritten As Integer, ByVal totalFilesCompleted As Long, ByVal totalFilesTotal As Long)

        ''' <summary>
        '''     <para>Fired when the decompression is completed</para>
        '''     <para>The object is automatically disposed (cannot be used anymore)</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="totalCompleted">Total bytes written</param>
        Public Event Finished(ByVal totalCompleted As ULong)

        ''' <summary>
        '''     <para>Fired when the process has been stopped</para>
        '''     <para>The object is automatically disposed (cannot be used anymore)</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        Public Event Aborted()

        'I/O
        Private inputFile As FileInfo
        Private outputDirectory As DirectoryInfo
        Private archive As ZipArchive
        'lengths
        Private totalLength As Long = 0
        Private completedLength As Long = 0
        'decompressor
        Private WithEvents decomp As Decompressor
        Private index As Integer = 0
        Private abortAll As Boolean = False

        ''' <summary>
        ''' Initializes a new ZIP archive decompressor
        ''' </summary>
        ''' <param name="inputArchive">ZIP archive path</param>
        ''' <param name="outputDir">Output directory for extracted files</param>
        Public Sub New(ByVal inputArchive As FileInfo, ByVal outputDir As DirectoryInfo)
            Me.inputFile = inputArchive
            Me.totalLength = inputArchive.Length
            Me.outputDirectory = outputDir
            Me.index = 0
            Me.archive = New ZipArchive(New FileStream(inputArchive.FullName, FileMode.Open, FileAccess.Read), ZipArchiveMode.Read)
        End Sub

        ''' <summary>
        ''' Starts the asynchronous decompression process
        ''' </summary>
        Public Sub Start()
            Me.decomp = New Decompressor(Me.inputFile, Me.archive.Entries.Item(Me.index).Name, New FileInfo(My.Computer.FileSystem.CombinePath(Me.outputDirectory.FullName, Me.archive.Entries.Item(Me.index).Name)))
            Me.decomp.StartDecompression()
        End Sub

        ''' <summary>
        ''' Stops the asynchronous process for all files and automatically dispose the object
        ''' </summary>
        Public Sub Cancel()
            Me.abortAll = True
            Me.decomp.Cancel()
        End Sub

        ''' <summary>
        ''' Flush, close and release all resources used by the decompressors
        ''' </summary>
        ''' See <see cref="IDisposable.Dispose()"/>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me.Cancel()
            Me.decomp.Dispose()
            Me.archive.Dispose()
        End Sub


        'decompressor events
        Private Sub decomp_Progress(bytesDecompressed As ULong, bytesTotal As ULong, nowWritten As UInteger) Handles decomp.Progress
            Me.completedLength += nowWritten
            RaiseEvent Progress(Me.index, bytesDecompressed, bytesTotal, nowWritten, Me.completedLength, Me.totalLength)
        End Sub

        Private Sub decomp_Finished(bytesWritten As ULong) Handles decomp.Finished
            If (Not Me.abortAll) Then
                Me.index += 1
                If (Me.index < Me.archive.Entries.Count) Then
                    Me.decomp = New Decompressor(Me.inputFile, Me.archive.Entries.Item(Me.index).Name, New FileInfo(My.Computer.FileSystem.CombinePath(Me.outputDirectory.FullName, Me.archive.Entries.Item(Me.index).Name)))
                    Me.decomp.StartDecompression()
                Else
                    RaiseEvent Finished(Me.completedLength)
                    Me.Dispose()
                End If
            Else
                RaiseEvent Aborted()
                Me.Dispose()
            End If
        End Sub

        Private Sub decomp_Aborted() Handles decomp.Aborted
            If (Not Me.abortAll) Then
                If (Me.index < Me.archive.Entries.Count) Then
                    Me.index += 1
                    Me.decomp = New Decompressor(Me.inputFile, Me.archive.Entries.Item(Me.index).Name, New FileInfo(My.Computer.FileSystem.CombinePath(Me.outputDirectory.FullName, Me.archive.Entries.Item(Me.index).Name)))
                    Me.decomp.StartDecompression()
                Else
                    RaiseEvent Finished(Me.completedLength)
                    Me.Dispose()
                End If
            Else
                RaiseEvent Aborted()
                Me.Dispose()
            End If
        End Sub
    End Class
End Namespace
