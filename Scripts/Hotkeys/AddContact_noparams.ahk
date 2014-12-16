; // Hotkey for Add Customer Contact
; // Created by: Zachary Karpinski
; // Last Modified: 12/16/14

; // usage: Autohotkey http://www.autohotkey.com/
; //	Enter account number, run script then make the CSS window active

; // Open file dialog with multiselect
FileSelectFile, SelectedFiles, M3, T:\Web Resources\SharedFTP\CollRep, Select the DPA files to add contacts to.
If SelectedFiles =
{
	MsgBox, No items selected.
	return
}

; // Get contact to add.
InputBox, strContact, Contact, Please enter a Contact.
if ErrorLevel
{
	MsgBox, CANCEL was pressed.
	return
}
else{}


Loop, parse, SelectedFiles, `n
{
	if a_index = 1 
	{}
	else
{
	; //StringSplit, splitArray, A_LoopField, -
	; //AccNum := splitArray2 . "-" . splitArray3
	RegExMatch(A_LoopField,"\d{5}-\d{5}",AccNum)
	; //RegExMatch(A_LoopField,(\b\d-\d(3)-\d{3}-\d{4} )

	
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
Send, %strContact%
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

}
}
