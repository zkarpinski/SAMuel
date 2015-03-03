; // Hotkey for Add Customer Contact(With Parameters)
; // Created by: Zachary Karpinski
; // Last Modified: 01/21/15

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

;Check there are values
If (AccNum == "") {
	Msgbox, No account number passed.
	Exit 88
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

; // Add Account number
Send {TAB 4}
Sleep, 250
Send, %AccNum%

; // Find Account number
Sleep, 250
Send, {ENTER}
Sleep, 1000

;Handle the premise window
IfWinExist, Premise for
{
	WinActivate
	Sleep, 50
	Send, {Alt}
	Sleep, 250
	Send, m
	Sleep, 100
	Send, e
	Sleep, 250
	Exit 4
}

; // Wait for Account (20 Sec)
WinWait, Account %AccNum% for,, 20
if ErrorLevel
{
	;Handle invalid account number
	IfWinExist, Customer Service System (PRODUCTION)
	{
		WinActivate
		Sleep, 100
		Send, o
		Exit 5
	}
	FullyCloseAccount(AccNum)
	Exit 1
}
Sleep, 100

; // Skip Critcal Contact message.
Send, o
Sleep, 100
Send, o
Sleep, 100

;// Close the Crtical Message (If it still exists)
IfWinExist, Customer Service System (PRODUCTION)
{
	WinActivate
	Sleep, 100
	Send, o
}


;// Focus The account Window
WinActivate, Account %AccNum% for
Sleep, 100

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

; // Wait for add account contact window (20 Sec)
WinWait, Add Account Contact for %AccNum%,, 20
if ErrorLevel
{
	FullyCloseAccount(AccNum)
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

; // Wait for confirmation window (20 seconds)
WinWait, Customer Service System (PRODUCTION),, 20
if ErrorLevel
{
	FullyCloseAccount(AccNum)
	Exit 3
}
Sleep, 50
Send, o
Sleep, 200

; // Exit Account
WinActivate, Account %AccNum% for
Sleep, 200
Send, {ALT}
Sleep, 100
Send, c
Sleep, 100
Send, e
Sleep 300

FullyCloseAccount(AccNum)

ExitApp 100


;Function to force close account
FullyCloseAccount(AccNum) {
	IfWinExist, Account %AccNum% for
	{
		WinActivate
		; // Exit Account
		Send, {ALT}
		Sleep, 100
		Send, c
		Sleep, 100
		Send, e
		Sleep 300
	}
}




