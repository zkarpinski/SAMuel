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

#Region "Deprecated Code. Save it for later."
    Private Const sWait As Short = 300

    'Define Keys
    Private kAlt As String = "%"
    Private kCtrl As String = "^"
    Private kDel As String = "{DEL}"
    Private kTab As String = "{TAB}"
    Private kUp As String = "{UP}"
    Private kSpace As String = "{SPACE}"
    Private kDown As String = "{DOWN}"
    Private kEnter As String = "~"

    Sub MiscCollection(accountNumber As String, strContact As String)
        modDPAHandle.OpenCSSAcc(accountNumber)
        Threading.Thread.Sleep(2000)
        OpenContactWindow()
        Threading.Thread.Sleep(1000)
        AddContact(strContact)
        Threading.Thread.Sleep(1000)
        CloseAccount()
    End Sub

    Private Sub OpenContactWindow()
        'Open Account Contact Window
        ' alt a, a, c, a
        SendKeys.Send(kAlt)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("a")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("a")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("c")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("a")
    End Sub

    Private Sub AddContact(strContact As String)
        'Add Contact to Account
        'Tab to Contact Type
        For i As Integer = 1 To 8 Step 1
            SendKeys.Send(kTab)
            Threading.Thread.Sleep(sWait)
        Next
        SendKeys.Send("Miscellaneous Collections")
        Threading.Thread.Sleep(800)
        SendKeys.Send(kTab)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(strContact)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kAlt)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("c")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("p")
        Threading.Thread.Sleep(1000)
        SendKeys.Send(kEnter)
    End Sub

    Private Sub CloseAccount()
        'Close the account window
        SendKeys.Send(kAlt)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("c")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("e")
    End Sub


#End Region

End Module
