; // Hotkey for Add Customer Contact(With Parameters)
; // Created by: Zachary Karpinski
; // Last Modified: 12/16/14

; // usage: Autohotkey http://www.autohotkey.com/
; //	Enter account number, run script then make the CSS window active

; // Exit Codes:
; // 0 Manually Exited
; // 100 Full Success

; // Read Parameters http://www.autohotkey.com/board/topic/6953-processing-command-line-parameters/
Loop, %0% {
    If (%A_Index% = "/account")	{ 
        AccNum := A_Index + 1
		AccNum := %AccNum%
		}
    Else If (%A_Index% = "/contact") {
        Contact := A_Index + 1
		Contact := %Contact%
		}
	Else If (%A_Index% = "/type") {
        ContactType := A_Index + 1
		ContactType := %ContactType%
		}
  }

; // Wait until CSS is loaded.
WinWait, Customer Service System Retrieval
WinActivate, Customer Service System Retrieval

; // Clear retriveval
Send, {ALT down}
Sleep, 50
Send, d
Sleep, 100
Send {ALT up}

; // Add premise number
Send {TAB}
Sleep, 250
Send, %AccNum%

; // Find account number
Sleep, 250
Send, {ENTER}
Sleep, 800

IfWinExist, Premise For
{
Sleep, 5000
}

; // Wait for Account
WinWait, Account ,, 120
if ErrorLevel
{
	MsgBox, Waiting for contact window timed out.
	Exit 1
}
else{}
Sleep, 50

; // Skip Critcal Contact message.
Send, o
Sleep, 250
Send, o

; // Open Contact Window
Send, {ALT}
Sleep, 250
Send, a
Sleep, 100
Send, a
Sleep, 100
Send, c
Sleep, 100
Send, a

; // Wait for add account contact window (1 minute)
WinWait, Add Account Contact for,, 60
if ErrorLevel
{
	MsgBox, Waiting for contact window timed out.
	Exit 2
}
Sleep, 100

; // Select Contact Type
Send, {TAB 8}
Sleep, 100
Send, %ContactType%
Sleep, 200

; //Enter Contact
Send, {TAB}
Sleep, 100
Send, %Contact%
Sleep, 200

; //Process
Send, {ALT}
Sleep, 100
Send, c
Sleep, 100
Send, p
; // Wait for confirmation window (10 seconds)
WinWait, Customer Service System (PRODUCTION),, 10
if ErrorLevel
{
	MsgBox, timed out.
	Exit 3
}
Sleep, 50
Send, o
Sleep, 200

; // Exit Account
Send, {ALT}
Sleep, 100
Send, c
Sleep, 100
Send, e
Sleep 300

ExitApp 100

