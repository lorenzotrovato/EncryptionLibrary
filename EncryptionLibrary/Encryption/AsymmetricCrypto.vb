Imports System.Text
Imports System.Threading
Imports System.Security.Cryptography


Namespace Encryption


    ''' <summary>
    ''' Provide methods for asynchronous asymmetric encryption and decryption of text and byte arrays
    ''' Note: asymmetric encryption is meant for small contents such as keys and passwords
    ''' </summary>
    ''' Base class <see cref="Crypto"/>
    Public Class AsymmetricCrypto
        Inherits Crypto
        Implements IDisposable

        'events
        ''' <summary>
        ''' Fired when the keys initialization process is completed and the object is usable
        ''' </summary>
        Public Event KeysGenerated()

        ''' <summary>
        '''     <para>Fired when the encryption or decryption process is completed</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="result">The result of the encryption or decryption</param>
        ''' <param name="isEncryption">Tell if is an encryption or decryption process</param>
        Public Event Finished(result As Byte(), isEncryption As Boolean)

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

        'input byte array
        Private inputByte() As Byte
        'output byte array
        Private outputByte() As Byte
        'cryptographic provider
        Private crProvider As RSACryptoServiceProvider
        'selected algorithm
        Private selectedAlgo As AsymmetricAlgos
        'hashed password
        Private hashedKey As Byte()
        'key generation thread
        Private threadKeyGen As Thread
        'tmp input
        Private tmpInput As Byte()

        'algorithms
        ''' <summary>
        ''' Available asymmetric key sizes for RSA algorithm
        ''' </summary>
        Public Enum AsymmetricAlgos As Integer
            RSA_Key16384bits = 16384
            RSA_Key8192bits = 8192
            RSA_Key4096bits = 4096
            RSA_Key2048bits = 2048
            RSA_Key1369bits = 1396
            RSA_Key1024bits = 1024
            RSA_Key952bits = 952
            RSA_Key512bits = 512
        End Enum

        ''' <summary>
        ''' Creates a new asymmetric encryptor
        ''' </summary>
        ''' <param name="input">The plain input. Input size should be less than MaxInputLength property, based on key size</param>
        ''' <param name="algo">The key size. Please note that bigger keys encrypt more data but are slower to generate</param>
        Public Sub New(ByVal input As Byte(), ByVal Optional algo As AsymmetricAlgos = 2048)
            MyBase.New()
            Me.selectedAlgo = algo
            Me.threadCrypto = New Thread(New ThreadStart(AddressOf AsyncCrypt))
            Me.threadDecrypto = New Thread(New ThreadStart(AddressOf AsyncDecrypt))
            Me.threadKeyGen = New Thread(New ThreadStart(AddressOf Keygen))
            Me.bufferLength = 0
            Me.tmpInput = input
            Me.threadKeyGen.Start()
        End Sub

        ''' <summary>
        ''' Creates a new asymmetric encryptor
        ''' </summary>
        ''' <param name="inputString">The plain input as string (will be converted to byte array). Input size should be less than MaxInputLength property, based on key size</param>
        ''' <param name="algo">The key size. Please note that bigger keys encrypt more data but are slower to generate</param>
        Public Sub New(inputString As String, ByVal Optional algo As AsymmetricAlgos = 2048)
            Me.New(Encoding.Unicode.GetBytes(inputString), algo)
        End Sub

        ''' <summary>
        ''' Starts the asynchronous encryption process
        ''' <para>Encryption process is one-time use. When the <c>Finished</c> event is raised, the objetc is automatically disposed</para>s
        ''' </summary>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the keys haven't been generated yet (KeysGenerated event)</exception>
        ''' See <see cref="Crypto.StartEncryption()"/>
        Public Overrides Sub StartEncryption()
            If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                Me.threadCrypto.Start()
            Else
                Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
            End If
        End Sub

        ''' <summary>
        ''' Starts the asynchronous decryption process, using the auto-generated keys
        ''' <para>Decryption process is one-time use. When the <c>Finished</c> event is raised, the objetc is automatically disposed</para>
        ''' </summary>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the keys haven't been generated yet (KeysGenerated event)</exception>
        ''' <see cref="Crypto.StartDecryption()"/>
        Public Overrides Sub StartDecryption()
            If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                Me.threadDecrypto.Start()
            Else
                Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
            End If
        End Sub

        ''' <summary>
        ''' Starts the asynchronous decryption process, importing the private key
        ''' <para>Decryption process is one-time use. When the <c>Finished</c> event is raised, the objetc is automatically disposed</para>
        ''' </summary>
        ''' <param name="keyPrivate">The private key to decrypt the data</param>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the object hasn't been initialized yet (KeysGenerated event)</exception>
        ''' <see cref="Crypto.StartDecryption()"/>
        Public Overloads Sub StartDecryption(ByVal keyPrivate As RSAParameters)
            If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                Me.crProvider = New RSACryptoServiceProvider(Me.selectedAlgo)
                Me.crProvider.ImportParameters(keyPrivate)
                Me.threadDecrypto.Start()
            Else
                Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
            End If
        End Sub

        ''' <summary>
        ''' Starts the asynchronous decryption process, importing the private key from XML string
        ''' <para>Decryption process is one-time use. When the <c>Finished</c> event is raised, the objetc is automatically disposed</para>
        ''' </summary>
        ''' <param name="XMLkeyPrivate">The private key to decrypt the data in XML format</param>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the object hasn't been initialized yet (KeysGenerated event)</exception>
        ''' See <see cref="Crypto.StartDecryption()"/>
        Public Overloads Sub StartDecryption(ByVal XMLkeyPrivate As String)
            If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                Me.crProvider = New RSACryptoServiceProvider(Me.selectedAlgo)
                Me.crProvider.FromXmlString(XMLkeyPrivate)
                Me.threadDecrypto.Start()
            Else
                Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
            End If
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
        ''' </summary>
        ''' See <see cref="Crypto.StopDecryption()"/>
        Public Overrides Sub StopDecryption()
            If (Me.threadDecrypto.ThreadState = ThreadState.Running) Then
                Me.threadDecrypto.Abort()
                Me.Dispose()
            End If
        End Sub

        ''' <summary>
        ''' Release all the resources used by the object
        ''' </summary>
        ''' See <see cref="IDisposable.Dispose()"/>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me.crProvider.Clear()
            Me.crProvider.Dispose()
        End Sub

        ''' <summary>
        ''' The public key used in asymmetric encryption. This could be revealed to others
        ''' </summary>
        ''' <returns>The public key as RSAParameters object</returns>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the object hasn't been initialized yet (KeysGenerated event)</exception>
        Public ReadOnly Property PublicKey As RSAParameters
            Get
                If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                    Return Me.crProvider.ExportParameters(False)
                Else
                    Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
                End If
            End Get
        End Property

        ''' <summary>
        ''' The private key used in asymmetric encryption. This should be kept secret
        ''' </summary>
        ''' <returns>The private key as RSAParameters object</returns>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the object hasn't been initialized yet (KeysGenerated event)</exception>
        Public ReadOnly Property PrivateKey As RSAParameters
            Get
                If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                    Return Me.crProvider.ExportParameters(True)
                Else
                    Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
                End If
            End Get
        End Property

        ''' <summary>
        ''' The private key used in asymmetric encryption. This should be kept secret
        ''' </summary>
        ''' <returns>The private key as XML formatted string</returns>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the object hasn't been initialized yet (KeysGenerated event)</exception>
        Public ReadOnly Property XMLPrivateKey As String
            Get
                If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                    Return Me.crProvider.ToXmlString(True)
                Else
                    Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
                End If
            End Get
        End Property

        ''' <summary>
        ''' The public key used in asymmetric encryption. This could be revealed to others
        ''' </summary>
        ''' <returns>The public key as XML formatted string</returns>
        ''' <exception cref="CryptographicUnexpectedOperationException">Thrown if the object hasn't been initialized yet (KeysGenerated event)</exception>
        Public ReadOnly Property XMLPublicKey As String
            Get
                If (TypeOf Me.crProvider Is RSACryptoServiceProvider And Me.bufferLength > 0) Then
                    Return Me.crProvider.ToXmlString(False)
                Else
                    Throw New CryptographicUnexpectedOperationException("Keys are not generated. Please wait until KeysGenerated event is fired")
                End If
            End Get
        End Property

        ''' <summary>
        ''' Maximum input length based on key size. If input is bigger, an <c>ArgumentOutOfRangeException</c> is thrown
        ''' </summary>
        ''' <returns>The maximum input length</returns>
        Public ReadOnly Property MaxInputLength As Integer
            Get
                Return Me.bufferLength
            End Get
        End Property

        'private methods
        Private Sub AsyncCrypt()
            Me.crProvider.PersistKeyInCsp = False
            Dim result As Byte() = Me.crProvider.Encrypt(Me.inputByte, True)
            Me.Dispose()
            RaiseEvent Finished(result, True)
        End Sub

        Private Sub AsyncDecrypt()
            Me.crProvider.PersistKeyInCsp = False
            Dim result As Byte() = Me.crProvider.Decrypt(Me.inputByte, True)
            Me.Dispose()
            RaiseEvent Finished(result, False)
        End Sub

        Private Sub AsymmetricCrypto_KeysGenerated() Handles Me.KeysGenerated
            If (Me.tmpInput.Length > Me.bufferLength) Then
                Throw New ArgumentOutOfRangeException("input", Me.tmpInput.Length, "Input should be shorter than MaxInputLength property (" & Me.bufferLength & " bytes)")
            Else
                Me.inputByte = New Byte(Me.tmpInput.Length - 1) {}
                Me.tmpInput.CopyTo(Me.inputByte, 0)
                Me.tmpInput = Nothing
            End If
        End Sub

        Private Sub Keygen()
            Me.crProvider = New RSACryptoServiceProvider(Me.selectedAlgo)
            Me.crProvider.PersistKeyInCsp = False
            Me.bufferLength = Me.crProvider.ExportParameters(False).Modulus.Length
            RaiseEvent KeysGenerated()
        End Sub
    End Class
End Namespace