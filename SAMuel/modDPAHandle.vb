Module modDPAHandle
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
        'Clear
        SendKeys.Send(kAlt + "{d}")
        SendKeys.Send(kEnter)
        'Enter account number into field
        SendKeys.Send("{TAB 4}")
        SendKeys.Send(accNum)
        'Open account
        SendKeys.Send(kEnter)
        'Handle random msgbox
        SendKeys.Send("o")
    End Sub

    Public Sub OpenPA()
        'Open Payment Agreement windows
        ' alt a, a, l, m
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
        SendKeys.Send(kEnter)
    End Sub

    Public Sub EnrollBB()
        'Enroll into budget billing
        'Open Maintain Program Agreement window
        SendKeys.Send(kAlt + "oag")

        'Select Budget Billing
        SendKeys.Send(kDown)
        SendKeys.Send(kEnter)

        'Select Budget Billing Request
        SendKeys.Send("{TAB 3}")
        SendKeys.Send(kEnter)

        'Enter BB remarks
        SendKeys.Send("{TAB 8}")
        SendKeys.Send(kEnter)
        SendKeys.Send("bb as per signed dpa")
        SendKeys.Send(kTab + kEnter)
        'Exit prompts
        SendKeys.Send("{TAB 3}")
        SendKeys.Send(kEnter)

        'Progress Actions
        ProcessAction()

        'Exit window
        SendKeys.Send(kAlt + "p")
        SendKeys.Send("e")
        'Close window
        SendKeys.Send(kAlt + "c")
        SendKeys.Send("e")

    End Sub

    Public Sub CloseWin()
        'Close window
        'alt, c, e
    End Sub

End Module
