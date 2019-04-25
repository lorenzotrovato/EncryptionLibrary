Imports System.Threading


Namespace Encryption


    ''' <summary>
    ''' Base abstract class for encryption implementations.
    ''' Use <c>SymmetricCrypto</c> or <c>AsymmetricCrypto</c> instead
    ''' </summary>
    Public MustInherit Class Crypto
        'events
        ''' <summary>
        '''     <para>Fired when the decryption process has thrown an Exception; usually is for wrong password</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="message">The message associated with the exception</param>
        Public Shadows Event DecryptorError(message As String)

        ''' <summary>
        '''     <para>Fired when a generic exception is thrown in other threads</para>
        '''     <para>WARNING: this event is raised on background thread so be careful on accessing controls</para>
        ''' </summary>
        ''' <param name="ex">The exception thrown</param>
        Public Event ExceptionThrown(ex As Exception)


        'working threads
        Protected threadCrypto As Thread
        Protected threadDecrypto As Thread
        'buffer length
        Protected bufferLength As Integer


        'base methods not implemented
        ''' <summary>
        ''' Starts the asynchronous encryption process
        ''' </summary>
        Public MustOverride Sub StartEncryption()

        ''' <summary>
        ''' Starts the asynchronous decryption process
        ''' </summary>
        Public MustOverride Sub StartDecryption()

        ''' <summary>
        ''' Stops the asynchronous encryption process
        ''' </summary>
        Public MustOverride Sub StopEncryption()

        ''' <summary>
        ''' Stops the asynchronous decryption process
        ''' </summary>
        Public MustOverride Sub StopDecryption()
    End Class
End Namespace
