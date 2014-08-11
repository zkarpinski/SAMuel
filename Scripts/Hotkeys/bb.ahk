; // Hotkey for Budget Billing Request
; // Created by: Zachary Karpinski

; // usage: Autohotkey http://www.autohotkey.com/
; //	Enter account number, run script then make the CSS window active

IfWinNotActive, Customer Service System Retrieval
WinWaitActive, Customer Service System Retrieval
Sleep, 500
Send, {ENTER}
Sleep, 500
Send, {ALTDOWN}{ALTUP}oag
Sleep, 500
Send, {DOWN}{ENTER}
Sleep, 500
Send, {TAB 3}{ENTER}
Sleep, 500
Send, {TAB 8}{ENTER}
Sleep, 500
Send, bb{SPACE}as{SPACE}per{SPACE}signed{SPACE}dpa
Sleep, 50
Send, {TAB}{ENTER}
Sleep, 100
Send, {TAB 3}{ENTER}
Sleep, 100
Send, {CTRLDOWN}p{CTRLUP}{ENTER}
Sleep, 100
Send, {ALTDOWN}{ALTUP}pe
Sleep, 100
Send, {ALTDOWN}{ALTUP}ce
