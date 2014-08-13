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
        MsgBox, 4, , The file is %A_LoopField%.  Continue?
        IfMsgBox, No, break
	; //RegExMatch(%A_LoopField%,(?P<AccNum>\d{5}-\d{5}))
	; //MsgBox AccNum
}
WinWait, Customer Service System Retrieval
WinActivate, Customer Service System Retrieval

Sleep, 500
Send, {ENTER}
Sleep, 500

; // Open Contact Window
Send, {ALT}
Sleep, 500
Send, a
Sleep, 200
Send, a
Sleep, 200
Send, c
Sleep, 200
Send, a

; // Wait for add account contact window (2 minutes)
WinWait, , Add Account Contact For, 120
if ErrorLevel
{
	MsgBox, Waiting for contact window timed out.
	return
}
Sleep, 100

; // Select Contact Type
Send, {TAB 8}
Sleep, 200
Send, Miscellaneous Collections
Sleep, 400

; //Enter Contact
Send, {TAB}
Sleep, 200
Send, faxed active dpa
Sleep, 400

; //Process
Send, {ALT}
Sleep, 200
Send, c
Sleep, 200
; //Send, p
