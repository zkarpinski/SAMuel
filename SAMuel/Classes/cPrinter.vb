Imports System.IO

'we want an abstract base class
Public MustInherit Class CPrinter

    Protected _strFileToPrint As String = ""

    Protected Property strFileToPrint() As String

        Set(ByVal value As String)
            _strFileToPrint = value
        End Set

        Get
            Return _strFileToPrint
        End Get

    End Property

    'save original settings
    Protected MustOverride Sub pushSettings()
    'restore original settings
    Protected MustOverride Sub popSettings()
    'start printing process
    Public Overridable Sub PrintFile()

        'Define properties for the print process
        Dim procStartInfo As ProcessStartInfo = Nothing

        If System.IO.File.Exists(_strFileToPrint) Then

            procStartInfo = New ProcessStartInfo
            procStartInfo.FileName = _strFileToPrint
            procStartInfo.Verb = "Print"

            'Make process invisible
            procStartInfo.CreateNoWindow = True
            procStartInfo.WindowStyle = ProcessWindowStyle.Hidden

            'start print process for the file with/from the associated application
            Dim procPrint As Process = Process.Start(procStartInfo)
            'give the system some time
            System.Threading.Thread.Sleep(2500)

            If Not procPrint Is Nothing Then
                procPrint.WaitForExit()
            End If

        End If

    End Sub

End Class
