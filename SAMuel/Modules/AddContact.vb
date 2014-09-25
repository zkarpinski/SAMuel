Namespace Modules
    Module AddContact

        Public Sub RunScript(strBillAccount As String, strContact As String)
            Dim arguments As String = String.Format(" /account {0} /contact ""{1}", strBillAccount, strContact)

            'Call the script and wait.
            Dim p As New Process
            Dim psi As New ProcessStartInfo(Application.StartupPath + "\addContact.exe", arguments)
            p.StartInfo = psi
            p.Start()
            p.WaitForExit()
        End Sub

    End Module
End Namespace