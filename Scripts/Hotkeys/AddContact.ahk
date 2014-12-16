; // Hotkey for Add Customer Contact(With Parameters)
; // Created by: Zachary Karpinski
; // Last Modified: 12/16/14

; // usage: Autohotkey http://www.autohotkey.com/
; //	Enter account number, run script then make the CSS window active

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

; // Add account number
Send {TAB 4}
Sleep, 250
Send, %AccNum%

; // Find account number
Sleep, 250
Send, {ENTER}
Sleep, 600

IfWinExist, Premise For
{
Sleep, 5000
}

; // Skip Critcal Contact message.
Send, o
Sleep, 250
Send, o

; // Wait for Account
WinWait, Account %AccNum% for,, 120
if ErrorLevel
{
	MsgBox, Waiting for contact window timed out.
	return
}
else{}
Sleep, 50


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
	return
}
Sleep, 100

; // Select Contact Type
Send, {TAB 8}
Sleep, 100
Send, Miscellaneous Collections
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
	return
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

