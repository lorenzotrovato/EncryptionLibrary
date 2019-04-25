# EncryptionLibrary
Easily compress and encrypt multiple files with several algorithms.
All the operations are on background threads.

## Encryption
### Symmetric key algorithms
Encrypt and decrypt files and byte streams with a simple password. This library uses the ``` System.Security.Cryptography ``` implementations and is programmed to load the system memory according to the file size and current available memory resulting in high performances and speed.
##### Available algorithms
* AES with SHA256 key (default)
* AES with MD5 key
* TripleDES with MD5 key
* RC2 with MD5 key
##### Events
* ``` Progress(bytesCompleted As Long, bytesTotal As Long, bufferLength As Integer) ```
* ``` Finished(bytesWritten As Long, isEncryption As Boolean) ```
* ``` DecryptorError(message As String) ``` (usually wrong password)
* ``` ExceptionThrown(ex As Exception) ``` (general errors)

### RSA asymmetric key algorithm
Encrypt and decrypt small texts, export and share public and private keys. 
##### Available keys sizes:
* 16384 bits
* 8192 bits
* 4096 bits
* 2048 bits (default)
* 1396 bits
* 1024 bits
* 952 bits
* 512 bits
##### Events
* ``` KeysGenerated() ```
* ``` Finished(result As Byte(), isEncryption As Boolean) ```
* ``` DecryptorError(message As String) ``` (usually wrong keys)
* ``` ExceptionThrown(ex As Exception) ``` (general errors)

### Utils
Calculate hashes from strings
##### Available hashing algorithms
* MD5
* SHA1
* SHA265
* SHA384
* SHA512

## Compression
Compress and decompress files with ZIP Deflate algorithm
* Compressor: compress single file
* Decompressor: decompress single file
* ArchiveCompressor: compress multiple files into a ZIP archive
* ArchiveDecompressor: decompress a ZIP archive into multiple files

## More info
Objects are automatically disposed after use.
Be careful on accessing graphical controls in events methods. They're all on separate threads so you should use delgate subs.
All classes are well documented. Look up there for more detailed usage informations.

### License
This library is released under GNU General Public License v2.0
Please include the license and quote the author when using it on public software.
