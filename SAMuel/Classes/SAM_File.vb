Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Namespace Classes
    Public Class SAM_File
        Private ReadOnly _file As String
        Private ReadOnly _filename As String
        Private ReadOnly _filesize As String
        Private ReadOnly _filetype As String
        Private ReadOnly _md5Hash As String

        Public Sub New(ByVal strFile As String)
            _file = strFile
            _filesize = 0
            _filename = Path.GetFileNameWithoutExtension(File)
            _filetype = Path.GetExtension(strFile)
            _md5Hash = HashFile(strFile)
        End Sub

        
        ''' <summary>
        '''     Computes the MD5 hash value of a file and returns the hash as a 32-character, hexadecimal-formatted string.
        ''' </summary>
        ''' <param name="input"></param>
        ''' <returns>Returns the hash as a 32-character, hexadecimal-formatted string. </returns>
        ''' <remarks>http://msdn.microsoft.com/en-us/library/system.security.cryptography.md5.aspx</remarks>
        Private Shared Function HashFile(ByVal input As String) As String

            Dim md5 As MD5 = New MD5CryptoServiceProvider()

            Using fs As FileStream = New FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
                md5.ComputeHash(fs)
            End Using

            Dim hash As Byte() = md5.Hash
            Dim buff As StringBuilder = New StringBuilder
            Dim hashByte As Byte

            For Each hashByte In hash
                buff.Append(String.Format("{0:X2}", hashByte))
            Next

            Dim md5String As String
            md5String = buff.ToString()

            Return md5String
        End Function

        Public ReadOnly Property File As String
            Get
                Return _file
            End Get
        End Property

        Public ReadOnly Property Filesize As String
            Get
                Return _filesize
            End Get
        End Property

        Public ReadOnly Property Filename As String
            Get
                Return _filename
            End Get
        End Property

        Public ReadOnly Property Filetype As String
            Get
                Return _filetype
            End Get
        End Property

        Public ReadOnly Property Md5Hash As String
            Get
                Return _md5Hash
            End Get
        End Property
    End Class
End Namespace