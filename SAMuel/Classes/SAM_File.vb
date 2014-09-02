Imports System.Text
Imports System.IO

Public Class SAM_File
    Private m_file As String
    Private m_filename As String
    Private m_filesize As String
    Private m_filetype As String
    Private m_md5hash As String

    Public Sub New(ByVal strFile As String)
        Me.m_file = strFile
        Me.m_filename = System.IO.Path.GetFileNameWithoutExtension(File)
        Me.m_filetype = System.IO.Path.GetExtension(strFile)
        Me.m_md5hash = HashFile(strFile)
    End Sub

    ''' <summary>
    ''' Computes the MD5 hash value of a file and returns the hash as a 32-character, hexadecimal-formatted string.
    ''' </summary>
    ''' <param name="input"></param>
    ''' <returns>Returns the hash as a 32-character, hexadecimal-formatted string. </returns>
    ''' <remarks>http://msdn.microsoft.com/en-us/library/system.security.cryptography.md5.aspx</remarks>
    Private Function HashFile(ByVal input As String) As String

        Dim md5 As System.Security.Cryptography.MD5 = New System.Security.Cryptography.MD5CryptoServiceProvider()

        Using fs As FileStream = New FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            md5.ComputeHash(fs)
        End Using

        Dim hash As Byte() = md5.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte

        For Each hashByte In hash
            buff.Append(String.Format("{0:X2}", hashByte))
        Next

        Dim md5string As String
        md5string = buff.ToString()

        Return md5string
    End Function

    Public ReadOnly Property File As String
        Get
            Return Me.m_file
        End Get
    End Property
    Public ReadOnly Property Filesize As String
        Get
            Return Me.m_filesize
        End Get
    End Property
    Public ReadOnly Property Filename As String
        Get
            Return Me.m_filename
        End Get
    End Property
    Public ReadOnly Property Filetype As String
        Get
            Return Me.m_filetype
        End Get
    End Property
    Public ReadOnly Property MD5Hash As String
        Get
            Return Me.m_md5hash
        End Get
    End Property

End Class
