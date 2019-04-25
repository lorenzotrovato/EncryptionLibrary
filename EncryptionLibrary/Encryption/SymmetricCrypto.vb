Imports System.IO
Imports System.Security.Cryptography
Imports System.Threading


Namespace Encryption


    ''' <summary>
    ''' Provide methods for asynchronous symmetric encryption and decryption of data streams such as files
    ''' </summary>
    ''' Base class <see cref="Crypto"/>
    Public Class SymmetricCrypto
        Inherits Crypto
        Implements IDisposable

        'events
        ''' <summary>
        '''     <para>Fired when a block of the stream has been encrypted or decrypted</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="bytesCompleted">Total bytes copied until now</param>
        ''' <param name="bytesTotal">Total length of the input stream</param>
        ''' <param name="bufferLength">Bytes read in this step</param>
        Public Event Progress(bytesCompleted As Long, bytesTotal As Long, bufferLength As Integer)

        ''' <summary>
        '''     <para>Fired when the encryption or decryption process is completed</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="bytesWritten">Total bytes written to the output stream</param>
        ''' <param name="isEncryption">Tell if is an encryption or decryption process</param>
        Public Event Finished(bytesWritten As Long, isEncryption As Boolean)

        ''' <summary>
        '''     <para>Fired when the decryption process has thrown an Exception; usually is for wrong password</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="message">The message associated with the exception</param>
        ''' See <see cref="Crypto.DecryptorError"/>
        Public Shadows Event DecryptorError(message As String)

        ''' <summary>
        '''     <para>Fired when a generic exception is thrown in other threads</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="ex">The exception thrown</param>
        ''' See <see cref="Crypto.ExceptionThrown"/>
        Public Shadows Event ExceptionThrown(ex As Exception)

        'input stream
        Protected inputStream As Stream
        'output stream
        Protected outputStream As Stream
        'input length completed
        Protected inputStreamCompleted As Long
        'crypto streams
        Private encryptionStream As CryptoStream
        Private decryptionStream As CryptoStream
        'cryptographic provider
        Private crProvider As SymmetricAlgorithm
        'selected algorithm
        Private selectedAlgo As SymmetricAlgos
        'hashed password
        Private hashedKey As Byte()


        'algorithms
        ''' <summary>
        ''' Available symmetric algorithms
        ''' </summary>
        Public Enum SymmetricAlgos As Integer
            AES_Key256bits = 0
            AES_Key128bits = 1
            TripleDES_Key_128bits = 2
            RC2_Key128bits = 3
        End Enum

        ''' <summary>
        ''' Creates a new symmetric encryptor
        ''' </summary>
        ''' <param name="password">The password used to encrypt or decrypt data</param>
        ''' <param name="inputStr">Input data stream</param>
        ''' <param name="outputStr">Output data stream</param>
        ''' <param name="algo">The algorithm used for encryption or decryption</param>
        ''' <param name="bufferLn">The buffer length used for stream copy (could cause OutOfMemoryException and switch to auto-mode</param>
        Public Sub New(ByVal password As String, ByVal inputStr As Stream, ByVal outputStr As Stream, ByVal Optional algo As SymmetricAlgos = 0, ByVal Optional bufferLn As Integer = 0)
            MyBase.New()
            Me.inputStream = inputStr
            Me.outputStream = outputStr
            Me.inputStreamCompleted = 0
            Me.bufferLength = Utils.CalculateCopyBufferLength(Me.inputStream.Length, bufferLn)
            Me.selectedAlgo = algo
            Select Case algo
                Case 0
                    Me.crProvider = Aes.Create()
                    Me.hashedKey = Utils.GetSHA256Array(password)
                Case 1
                    Me.crProvider = Aes.Create()
                    Me.hashedKey = Utils.GetMD5Array(password)
                Case 2
                    Me.crProvider = RC2.Create()
                    Me.hashedKey = Utils.GetMD5Array(password)
                Case 3
                    Me.crProvider = TripleDES.Create()
                    Me.hashedKey = Utils.GetMD5Array(password)
            End Select
            Me.threadCrypto = New Thread(New ThreadStart(AddressOf AsyncCrypt))
            Me.threadDecrypto = New Thread(New ThreadStart(AddressOf AsyncDecrypt))
        End Sub

        ''' <summary>
        ''' Creates a new symmetric encryptor
        ''' </summary>
        ''' <param name="password">The password used to encrypt or decrypt data</param>
        ''' <param name="inputFilePath">Path of the input file</param>
        ''' <param name="outputFilePath">Path of the output file (must have write access)</param>
        ''' <param name="algo">The algorithm used for encryption or decryption</param>
        ''' <param name="bufferLn">The buffer length used for stream copy (could cause OutOfMemoryException and switch to auto-mode</param>
        Public Sub New(ByVal password As String, inputFilePath As String, outputFilePath As String, ByVal Optional algo As SymmetricAlgos = 0, Optional bufferLn As Integer = 0)
            Me.New(password, New FileStream(inputFilePath, FileMode.Open, FileAccess.Read), New FileStream(outputFilePath, FileMode.Create, FileAccess.Write), algo, bufferLn)
        End Sub

        ''' <summary>
        ''' Starts the asynchronous encryption process
        ''' <para>Encryption process is one-time use. When the <c>Finished</c> event is raised, the objetc is automatically disposed</para>
        ''' </summary>
        ''' See <see cref="Crypto.StartEncryption()"/>
        Public Overrides Sub StartEncryption()
            Me.encryptionStream = New CryptoStream(Me.outputStream, Me.crProvider.CreateEncryptor(Me.hashedKey, Me.crProvider.IV), CryptoStreamMode.Write)
            Me.encryptionStream.Write(Me.crProvider.IV, 0, Me.crProvider.IV.Length)
            Me.inputStreamCompleted = 0
            Me.threadCrypto.Start()
        End Sub

        ''' <summary>
        ''' Stops the asynchronous encryption process 
        ''' </summary>
        ''' See <see cref="Crypto.StopEncryption()"/>
        Public Overrides Sub StopEncryption()
            If (Me.threadCrypto.ThreadState = ThreadState.Running) Then
                Me.threadCrypto.Abort()
                Me.Dispose()
            End If
        End Sub

        ''' <summary>
        ''' Starts the asynchronous decryption process
        ''' <para>Decryption process is one-time use. When the <c>Finished</c> event is raised, the objetc is automatically disposed</para>
        ''' </summary>
        ''' See <see cref="Crypto.StartDecryption()"/>
        Public Overrides Sub StartDecryption()
            Dim iv(15) As Byte
            Me.inputStream.Read(iv, 0, 16)
            Me.decryptionStream = New CryptoStream(Me.inputStream, Me.crProvider.CreateDecryptor(Me.hashedKey, iv), CryptoStreamMode.Read)
            Me.inputStreamCompleted = 0
            Me.threadDecrypto.Start()
        End Sub

        ''' <summary>
        ''' Stops the asynchronous decryption process
        ''' </summary>
        ''' See <see cref="Crypto.StopDecryption()"/>
        Public Overrides Sub StopDecryption()
            If (Me.threadDecrypto.ThreadState = ThreadState.Running) Then
                Me.threadDecrypto.Abort()
                Me.Dispose()
            End If
        End Sub

        ''' <summary>
        ''' Flush, close and release all resources used by the streams
        ''' </summary>
        ''' See <see cref="IDisposable.Dispose()"/>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Me.inputStream.Flush()
                Me.outputStream.Flush()
                Me.inputStream.Close()
                Me.inputStream.Close()
                If (TypeOf Me.encryptionStream Is CryptoStream) Then
                    Me.encryptionStream.Flush()
                    Me.encryptionStream.Close()
                    Me.encryptionStream.Clear()
                ElseIf (TypeOf Me.decryptionStream Is CryptoStream) Then
                    Me.decryptionStream.Flush()
                    Me.decryptionStream.Close()
                    Me.decryptionStream.Clear()
                End If
                Me.inputStream.Dispose()
                Me.outputStream.Dispose()
            Catch ex As Exception

            End Try
        End Sub

        'private methods
        Private Sub AsyncCrypt()
            Dim current() As Byte
            Try
                current = New Byte(Me.bufferLength) {}
            Catch ex As OutOfMemoryException
                'switch to auto managing
                Me.bufferLength = Utils.CalculateCopyBufferLength(Me.inputStream.Length)
                current = New Byte(Me.bufferLength) {}
                RaiseEvent ExceptionThrown(ex)
            End Try
            Dim currentRead As Long = 0
            While (Me.inputStreamCompleted < Me.inputStream.Length)
                Try
                    currentRead = Me.inputStream.Read(current, 0, Me.bufferLength)
                    Me.encryptionStream.Write(current, 0, currentRead)
                    Me.inputStreamCompleted = Me.inputStream.Position
                    RaiseEvent Progress(Me.inputStreamCompleted, Me.inputStream.Length, currentRead)
                Catch e As OutOfMemoryException
                    'switch to auto managing
                    Me.bufferLength = Utils.CalculateCopyBufferLength(Me.inputStream.Length)
                    RaiseEvent ExceptionThrown(e)
                Catch ex As Exception
                    RaiseEvent ExceptionThrown(ex)
                End Try
            End While
            RaiseEvent Progress(Me.inputStreamCompleted, Me.inputStream.Length, currentRead)
            Me.Dispose()
            RaiseEvent Finished(Me.inputStreamCompleted, True)
        End Sub

        Private Sub AsyncDecrypt()
            Dim current() As Byte
            Try
                current = New Byte(Me.bufferLength) {}
            Catch ex As OutOfMemoryException
                'switch to auto managing
                Me.bufferLength = Utils.CalculateCopyBufferLength(Me.inputStream.Length)
                current = New Byte(Me.bufferLength) {}
                RaiseEvent ExceptionThrown(ex)
            End Try
            Dim currentRead As Long = 0
            While (Me.inputStreamCompleted < Me.inputStream.Length)
                Try
                    currentRead = Me.decryptionStream.Read(current, 0, Me.bufferLength)
                    Me.outputStream.Write(current, 0, currentRead)
                    Me.inputStreamCompleted = Me.inputStream.Position
                    RaiseEvent Progress(Me.inputStreamCompleted, Me.inputStream.Length, currentRead)
                Catch ex As CryptographicException
                    RaiseEvent DecryptorError("Wrong password or damaged file")
                Catch ex1 As Exception
                    RaiseEvent ExceptionThrown(ex1)
                End Try
            End While
            RaiseEvent Progress(Me.inputStreamCompleted, Me.inputStream.Length, currentRead)
            Me.Dispose()
            RaiseEvent Finished(Me.inputStreamCompleted, False)
        End Sub
    End Class
End Namespace
