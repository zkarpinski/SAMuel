; // Hotkey for Opening the Account in CSS
; // Created by: Zachary Karpinski
; // Last Modified: 02/25/15

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