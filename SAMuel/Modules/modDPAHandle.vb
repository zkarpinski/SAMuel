Module modDPAHandle
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


    Public Sub OpenCSSAcc(ByVal accNum As String)
        'Focus to CSS
        AppActivate("Customer Service System Retrieval")
        Threading.Thread.Sleep(sWait)
        'Clear
        SendKeys.Send(kAlt + "{d}")
        Threading.Thread.Sleep(sWait)
        'Enter account number into field
        For i As Integer = 1 To 4 Step 1
            SendKeys.Send(kTab)
            Threading.Thread.Sleep(sWait)
        Next
        SendKeys.Send(accNum)
        Threading.Thread.Sleep(sWait)
        'Open account
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(700)
        'Handle random msgbox
        'SendKeys.Send("o")
        'Threading.Thread.Sleep(sWait)
    End Sub

    Public Sub OpenPA()
        'Open Payment Agreement windows
        ' alt a, a, l, m
        SendKeys.Send(kAlt)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("a")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("a")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("l")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("m")
        Threading.Thread.Sleep(sWait)


    End Sub

    Public Sub AddPA()
        'Add payment agreement (pending)
        'Click add

        'Negociated std

        'installment

        'down payment

        'check don't send

        'enter customer contact
    End Sub

    Public Sub ProcessAction()
        'Process Action
        'ctrl+p, enter
        SendKeys.Send(kCtrl + "{p}")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(400)
    End Sub

    Public Sub EnrollBB()
        Dim sBBKeys() As String = {kAlt, "o", "a", "g", _
                    kDown, kEnter, _
                    "{TAB 3}", kEnter, _
                    "{Tab 8}", kEnter, _
                    "bb as per signed dpa", _
                    kTab, kEnter _
                    }

        'Enroll into budget billing
        'Open Maintain Program Agreement window
        SendKeys.Send(kAlt)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("o")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("a")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("g")
        Threading.Thread.Sleep(400)

        'Select Budget Billing
        SendKeys.Send(kDown)
        Threading.Thread.Sleep(50)
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(400)

        'Select Budget Billing Request
        SendKeys.Send("{TAB 3}")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(400)

        'Enter BB remarks
        SendKeys.Send("{TAB 8}")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(400)
        SendKeys.Send("bb as per signed dpa")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kTab)
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(400)
        'Exit prompts
        SendKeys.Send("{TAB 3}")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send(kEnter)
        Threading.Thread.Sleep(400)

        'Progress Actions
        ProcessAction()

        'Exit window
        SendKeys.Send(kAlt + "p")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("e")
        Threading.Thread.Sleep(400)
        'Close window
        CloseWin()

    End Sub

    Private Sub CloseWin()
        'Close window
        'alt, c, e
        SendKeys.Send(kAlt + "c")
        Threading.Thread.Sleep(sWait)
        SendKeys.Send("e")
        Threading.Thread.Sleep(400)
    End Sub

End Module
