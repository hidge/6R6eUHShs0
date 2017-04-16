#SingleInstance, Force
#NoEnv
#Persistent
AutoTrim, On
SetWorkingDir %A_ScriptDir%\

;
;FINDATTACHED
;
+NumpadEnd::

objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
To := objOL.To

data =
(
%To%
)
loop, parse, data, `n,`r
FirstName(A_LoopField)
return
 
FirstName(_Name){
if !RegExMatch(_Name, "(\w+)\s\w+", recipient)
if !RegExMatch(_Name, "\w+`,\s(\w+)", recipient)
if !RegExMatch(_Name, "(\w+)\.\w+@", recipient)
recipient1 := _Name

objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
Sub := objOL.Subject
StringGetPos, pos, Sub, :, 1
Subject := SubStr(Sub, pos+2)

SendInput,Hi %recipient1%,
SendInput,{Enter}
SendInput,{Enter}
SendInput,Hope you are well.
SendInput,{Enter}
SendInput,{Enter}
SendInput,Find attached for the
Send,{Space}
Send,^b
SendInput,%Subject%
Send,^b
Send,{Space}
SendInput, service. Please confirm dates are viable and access can be granted.
SendInput,{Enter}
SendInput,{Enter}
SendInput,{Enter}
SendInput,Kind regards,
return
}


;
;CONFIRM
;

+NumpadDOWN::

objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
To := objOL.To

data =
(
%To%
)
loop, parse, data, `n,`r
FirstName1(A_LoopField)
return
 
FirstName1(_Name){
if !RegExMatch(_Name, "(\w+)\s\w+", recipient)
if !RegExMatch(_Name, "\w+`,\s(\w+)", recipient)
if !RegExMatch(_Name, "(\w+)\.\w+@", recipient)
recipient2 := _Name


SendInput,Hi %recipient1%,
SendInput,{Enter}
SendInput,{Enter}
SendInput,Thank you for confirming.
SendInput,{Enter}
SendInput,{Enter}
SendInput,{Enter}
SendInput,Kind regards,
return
}


;
;PUSHED
;

+NumpadPGDN::

; objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
To := objOL.To

data =
(
%To%
)
loop, parse, data, `n,`r
FirstName2(A_LoopField)
return
 
FirstName2(_Name){
if !RegExMatch(_Name, "(\w+)\s\w+", recipient)
if !RegExMatch(_Name, "\w+`,\s(\w+)", recipient)
if !RegExMatch(_Name, "(\w+)\.\w+@", recipient)
recipient1 := _Name

objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
Sub := objOL.Subject
StringGetPos, pos, Sub, :, 1
Subject := SubStr(Sub, pos+3)


SendInput,Hi %recipient1%,
SendInput,{Enter}
SendInput,{Enter}
SendInput,Sincere apologies, but due to unforeseen circumstances we are required to push the
Send,{Space}
Send,^b
SendInput,%Subject%
Send,^b
Send,{Space}
SendInput,service. We will reschedule for ASAP and come back to you shortly.
SendInput,{Enter}
SendInput,{Enter}
SendInput,Apologies for any inconveniences this may cause.
SendInput,{Enter}
SendInput,{Enter}
SendInput,{Enter}
SendInput,Kind regards,
return
}


;
;QUOTE
;

+NumpadLEFT::

; objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
To := objOL.To

data =
(
%To%
)
loop, parse, data, `n,`r
FirstName3(A_LoopField)
return
 
FirstName3(_Name){
if !RegExMatch(_Name, "(\w+)\s\w+", recipient)
if !RegExMatch(_Name, "\w+`,\s(\w+)", recipient)
if !RegExMatch(_Name, "(\w+)\.\w+@", recipient)
recipient1 := _Name

objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
Sub := objOL.Subject
StringGetPos, pos, Sub, :, 1
Subject := SubStr(Sub, pos+3)


SendInput,Hi %recipient1%,
SendInput,{Enter}
SendInput,{Enter}
SendInput,Please find attached our quotation for the
Send,{Space}
Send,^b
SendInput,%Subject%
Send,^b
Send,{Space}
SendInput,service.
SendInput,{Enter}
SendInput,{Enter}
SendInput,I look forward to receiving your further instructions in due course. 
SendInput,{Enter}
SendInput,{Enter}
SendInput,{Enter}
SendInput,Kind regards,
return
}


;
;QUOTE
;

+NumpadClear::

; objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
To := objOL.To

data =
(
%To%
)
loop, parse, data, `n,`r
FirstName4(A_LoopField)
return
 
FirstName4(_Name){
if !RegExMatch(_Name, "(\w+)\s\w+", recipient)
if !RegExMatch(_Name, "\w+`,\s(\w+)", recipient)
if !RegExMatch(_Name, "(\w+)\.\w+@", recipient)
recipient1 := _Name

objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
Sub := objOL.Subject
StringGetPos, pos, Sub, :, 1
Subject := SubStr(Sub, pos+3)


SendInput,Hi %recipient1%,
SendInput,{Enter}
SendInput,{Enter}
SendInput,Please find attached amended paperwork for the
Send,{Space}
Send,^b
SendInput,%Subject%
Send,^b
Send,{Space}
Send,service.
SendInput,{Enter}
SendInput,{Enter}
SendInput,{Enter}
SendInput,Kind regards,
return
}

f12::reload