Imports System.Text
Imports System.Security


Namespace Encryption


    ''' <summary>
    ''' Contains useful methods for hashing
    ''' </summary>
    Public Module Utils
        'HASH CALCULATORS

        '32 characters - 128 bits
        ''' <summary>
        ''' Calculate MD5 hash for the string given (32 characters, 128 bits)
        ''' </summary>
        ''' <param name="str">Plain string</param>
        ''' <returns>The hash of the string given</returns>
        Public Function GetMD5Array(ByVal str As String) As Byte()
            Dim sh As New Cryptography.MD5CryptoServiceProvider()
            Dim strbytes() As Byte = Encoding.UTF8.GetBytes(str)
            Dim sh1 As New Cryptography.SHA1Managed
            Return sh.ComputeHash(strbytes)
        End Function

        '40 characters - 160 bits
        ''' <summary>
        ''' Calculate SHA1 hash for the string given (40 characters, 160 bits)
        ''' </summary>
        ''' <param name="str">Plain string</param>
        ''' <returns>The hash of the string given</returns>
        Public Function GetSHA1Array(ByVal str As String) As Byte()
            Dim sh As New Cryptography.SHA1Managed()
            Dim strbytes() As Byte = Encoding.UTF8.GetBytes(str)
            Return sh.ComputeHash(strbytes)
        End Function

        '64 characters - 256 bits
        ''' <summary>
        ''' Calculate SHA256 hash for the string given (64 characters, 256 bits)
        ''' </summary>
        ''' <param name="str">Plain string</param>
        ''' <returns>The hash of the string given</returns>
        Public Function GetSHA256Array(ByVal str As String) As Byte()
            Dim sh As New Cryptography.SHA256Managed()
            Dim strbytes As Byte() = Encoding.UTF8.GetBytes(str)
            Return sh.ComputeHash(strbytes)
        End Function

        '96 characters - 384 bits
        ''' <summary>
        ''' Calculate SHA384 hash for the string given (96 characters, 384 bits)
        ''' </summary>
        ''' <param name="str">Plain string</param>
        ''' <returns>The hash of the string given</returns>
        Public Function GetSHA384Array(ByVal str As String) As Byte()
            Dim sh As New Cryptography.SHA384Managed()
            Dim strbytes() As Byte = Encoding.UTF8.GetBytes(str)
            Return sh.ComputeHash(strbytes)
        End Function

        '128 characters - 512 bits
        ''' <summary>
        ''' Calculate SHA512 hash for the string given (128 characters, 512 bits)
        ''' </summary>
        ''' <param name="str">Plain string</param>
        ''' <returns>The hash of the string given</returns>
        Public Function GetSHA512Array(ByVal str As String) As Byte()
            Dim sh As New Cryptography.SHA512Managed()
            Dim strbytes As Byte() = Encoding.UTF8.GetBytes(str)
            Return sh.ComputeHash(strbytes)
        End Function


        'OTHERS

        'buffer length auto based on ram
        ''' <summary>
        ''' Calculate the optimal size of the buffer on encryption and decryption processes
        ''' </summary>
        ''' <param name="streamlength">The total length of the stream</param>
        ''' <param name="preferredLength">Override the size (must be under 2GB for 32 bit process or under the current available free RAM for 64 bit process)</param>
        ''' <returns>The calculated size if preferredLength is 0 or preferredLength value</returns>
        Public Function CalculateCopyBufferLength(ByVal streamlength As ULong, ByVal Optional preferredLength As ULong = 0) As ULong
            If (preferredLength > 0) Then
                Return preferredLength
            Else
                If (streamlength > 2000) Then
                    Return streamlength / 1000
                Else
                    Return 1000
                End If
            End If
        End Function


        Public Function GetEnumValue(ByVal v As SymmetricCrypto.SymmetricAlgos) As Integer
            Return v
        End Function

        Public Function GetEnumValue(ByVal v As AsymmetricCrypto.AsymmetricAlgos) As Integer
            Return v
        End Function
    End Module
End Namespace
