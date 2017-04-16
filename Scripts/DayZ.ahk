#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#IfWinActive, DayZ

numpadup::
SprintToggle := !SprintToggle
If (SprintToggle)
{ Send, {lshift Down}
  Send, {w Down}
}
else
{ Send, {lshift up}
  Send, {w up}
}
return


NUMPADINS::
ReloadToggle := !ReloadToggle
If (ReloadToggle)
loop, 2000
{
send {r}
sleep 1
}
return