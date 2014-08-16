; // Hotkey for Add Customer Contact
; // Created by: Zachary Karpinski

; // usage: Autohotkey http://www.autohotkey.com/
; //	Enter account number, run script then make the CSS window active

; // Open file dialog with multiselect
FileSelectFile, SelectedFiles, M3
If SelectedFiles =
{
	MsgBox, No items selected.
	return
}
Loop, parse, SelectedFiles, `n
{
	if a_index = 1 
	{}
	else
{
	StringSplit, splitArray, A_LoopField, -
	AccNum := splitArray2 . splitArray3
	; //RegExMatch(A_LoopField,(?P<AccNum>\d{5}-\d{5}))
	; //  MsgBox %AccNum%
WinWait, Customer Service System Retrieval
WinActivate, Customer Service System Retrieval

Send, {ALT down}
Sleep, 50
Send, d
Sleep, 100
Send {ALT up}

Send {TAB 4}
Sleep, 250
Send, %AccNum%


Sleep, 250
Send, {ENTER}
Sleep, 600

IfWinExist, Premise For
{
Sleep, 5000
}

Send, o
Sleep, 250
Send, o
Sleep, 500

; // Wait for Account
; //WinWait, Account, 120


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

; // Wait for add account contact window (2 minutes)
;//WinWait, Add Account Contact for, 120
if ErrorLevel
{
	MsgBox, Waiting for contact window timed out.
	return
}
Sleep, 500

; // Select Contact Type
Send, {TAB 8}
Sleep, 100
Send, Miscellaneous Collections
Sleep, 200

; //Enter Contact
Send, {TAB}
Sleep, 100
Send, emailed active dpa
Sleep, 200

; //Process
Send, {ALT}
Sleep, 100
Send, c
Sleep, 100
Send, p
Sleep, 200
Send, o
Sleep, 150

; // Exit Account
Send, {ALT}
Sleep, 100
Send, c
Sleep, 100
Send, e
Sleep 300

}
}

