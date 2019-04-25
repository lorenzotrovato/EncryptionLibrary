Imports System.IO

Namespace Compression.Archive

    ''' <summary>
    ''' Creates or updates a ZIP archive
    ''' </summary>
    Public Class ArchiveCompressor
        Implements IDisposable

        ''' <summary>
        '''     <para>Fired when a block of one of the streams has been compressed</para>
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
        '''     <para>Fired when the compression is completed</para>
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
        Private inputFiles As List(Of FileInfo)
        Private outputFile As FileInfo
        'lengths
        Private totalLength As Long = 0
        Private completedLength As Long = 0
        'compressor
        Private WithEvents comp As Compressor
        Private index As Integer = 0
        Private abortAll As Boolean = False

        ''' <summary>
        ''' Initializes a new compressor for files
        ''' </summary>
        ''' <param name="inputs">A list of files to be compressed</param>
        ''' <param name="outputArchive">Path of the output archive</param>
        Public Sub New(ByVal inputs As List(Of FileInfo), ByVal outputArchive As FileInfo)
            Me.outputFile = outputArchive
            Me.inputFiles = New List(Of FileInfo)
            For Each el In inputs
                Me.inputFiles.Add(el)
                Me.totalLength += el.Length
            Next
            Me.index = 0
            Me.completedLength = 0
        End Sub

        ''' <summary>
        ''' Starts the asynchronous compression process
        ''' </summary>
        Public Sub Start()
            Me.comp = New Compressor(Me.inputFiles(Me.index), Me.outputFile)
            Me.comp.StartCompression()
        End Sub

        ''' <summary>
        ''' Stops the asynchronous process for all files and automatically dispose the object
        ''' </summary>
        Public Sub Cancel()
            Me.abortAll = True
            Me.comp.Cancel()
        End Sub

        ''' <summary>
        ''' Flush, close and release all resources used by the compressors
        ''' </summary>
        ''' See <see cref="IDisposable.Dispose()"/>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me.Cancel()
            Me.comp.Dispose()
        End Sub


        'compressor events
        Private Sub comp_Progress(bytesCompressed As ULong, bytesTotal As ULong, nowWritten As UInteger) Handles comp.Progress
            Me.completedLength += nowWritten
            RaiseEvent Progress(Me.index, bytesCompressed, bytesTotal, nowWritten, Me.completedLength, Me.totalLength)
        End Sub

        Private Sub comp_Finished(bytesWritten As ULong) Handles comp.Finished
            If (Not Me.abortAll) Then
                Me.index += 1
                If (Me.index < Me.inputFiles.Count) Then
                    Me.comp = New Compressor(Me.inputFiles(Me.index), Me.outputFile)
                    Me.comp.StartCompression()
                Else
                    RaiseEvent Finished(Me.completedLength)
                    Me.Dispose()
                End If
            Else
                RaiseEvent Aborted()
                Me.Dispose()
            End If
        End Sub

        Private Sub comp_Aborted() Handles comp.Aborted
            If (Not Me.abortAll) Then
                If (Me.index < Me.inputFiles.Count) Then
                    Me.index += 1
                    Me.comp = New Compressor(Me.inputFiles(Me.index), Me.outputFile)
                    Me.comp.StartCompression()
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

