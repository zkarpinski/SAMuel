
Imports Microsoft.Win32
Imports System.IO

Public Class cPDF995 : Inherits CPrinter

    'the INI that we want to change in oder to "silently" use the PDF printer
    Private Const strOriginalIniFile As String = "pdf995.ini"
    'we want to recover the original INI file after every print
    Private Const strOriginalIniSave As String = "pdf995.ini_ORIGINAL"
    'used to locate PDF995 installation directory
    Private Const strRegPath As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall\Pdf995"

    Private Function RetrieveIniPath() As String

        Dim strPdf995IniPath As String = ""

        Dim regKey As RegistryKey = Nothing
        Dim regKeyPDF995 As RegistryKey = Nothing

        regKey = Registry.LocalMachine.OpenSubKey(strRegPath)

        'retrieve path using the "DisplayIcon" value
        strPdf995IniPath = regKey.GetValue("DisplayIcon", "")

        'icon is in pdf995path\setup.exe and INI-File is in pdf995path\res directory, so remove "setup.exe"
        strPdf995IniPath = Strings.Left(strPdf995IniPath, strPdf995IniPath.Length - 9) + "res\"

        regKey.Close()

        Return strPdf995IniPath

    End Function

    Property fileToPrint() As String
        Get
            Return MyBase.strFileToPrint
        End Get
        Set(ByVal value As String)
            MyBase.strFileToPrint = value
        End Set
    End Property

    Protected Overrides Sub pushSettings()

        Dim strIniPath As String = RetrieveIniPath()

        'save original INI file first
        'File.Move(strIniPath + strOriginalIniFile, strIniPath + strOriginalIniSave)

        createCustomIni()

    End Sub

    Protected Overrides Sub popSettings()

        Dim strIniPath As String = RetrieveIniPath()

        'restore original INI file
        File.Delete(strIniPath + strOriginalIniFile)
        FileSystem.Rename(strIniPath + strOriginalIniSave, strIniPath + strOriginalIniFile)

    End Sub

    Private Sub createCustomIni()

        Dim strFileToPrint As String = MyBase.strFileToPrint
        Dim strIniPath As String = RetrieveIniPath()
        Dim fi As System.IO.FileInfo = New System.IO.FileInfo(strFileToPrint)
        Dim strPDFOutputDirectory As String = fi.DirectoryName + "\"
        Dim oWriter As StreamWriter = Nothing

        fi = Nothing

        'write INI file, create PDF from "strFileToPrint" and save it to "directory"
        oWriter = New StreamWriter(strIniPath + "pdf995.ini", False, System.Text.Encoding.Unicode)

        oWriter.WriteLine("[Parameters]")
        oWriter.WriteLine("Install = 0")
        oWriter.WriteLine("Quiet = 1")
        oWriter.WriteLine("AutoLaunch = 0")
        oWriter.WriteLine("Document Name = " + strFileToPrint)
        oWriter.WriteLine("User File = " + strFileToPrint + ".pdf")
        oWriter.WriteLine("Output File = " + strFileToPrint + ".pdf")
        oWriter.WriteLine("Initial Dir = " + strPDFOutputDirectory)
        oWriter.WriteLine("Use GPL Ghostcript = 1")

        oWriter.Close()

        oWriter = Nothing

        System.Threading.Thread.Sleep(1000) ' give Pdf995 some time

    End Sub

    Public Overrides Sub printFile()

        Me.pushSettings()
        MyBase.PrintFile()
        Me.popSettings()

    End Sub

End Class
