#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook ; to detect pressed keys
#InstallMouseHook ; to detect clicks
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Menu, Tray, Icon, Lcuts.ico ;Load icon
Menu, Tray, Tip, Lcuts ;Change display name

;Set Lock keys permanently
SetNumlockState, AlwaysOn
SetCapsLockState, AlwaysOff
SetScrollLockState, AlwaysOff

Return

^#F8::SwitchStartP("outlook.exe", True)
^#F9::SwitchStartP("slack.exe")
^#F10::SwitchStartP("C:\Program Files\Altap Salamander\salamand.exe")
^#F11::SwitchStartP("excel.exe")
^#F12::SwitchStartP("chrome.exe")

; XButton1::
; ;^#left
	; Send {Ctrl down}{LWin down}{Left down}
	; Sleep 50
	; Send {Left up}{LWin up}{Ctrl up}
; Return
; XButton2::
; ;^#right
	; Send {Ctrl down}{LWin down}{Right down}
	; Sleep 50
	; Send {Right up}{LWin up}{Ctrl up}
; Return

;===================
;=      WORD       =
;===================
;NumpadDiv::SwitchStartP("winword.exe")

;===================
;=      EXCEL      =
;===================
;NumpadMult::SwitchStartP("excel.exe")

;===================
;=     ALISTA      =
;===================
;NumpadAdd::Run, "C:\Program Files\internet explorer\iexplore.exe" "http://msdn.microsoft.com/en-us/library/aa394372.aspx"

;===================
;=   WINDOW TITLE  =
;===================
;Shows the title of the window being hovered on a traytip
; NumpadSub::
; {
	; MouseGetPos, , , id, control ;Gets the id of the window being hovered
	; WinGetTitle, Title, ahk_id %id%	;Get the title of the window from its id
	; Traytip, Window, %Title%, 0.5, 1
	; Return
; }

;===================
;=    MINIMIZE     =
;===================
;Minimizes the window being hovered
; NumpadDot::
; {
	; MouseGetPos, , , id, control ;Gets the id of the window being hovered
	; WinGetTitle, Title, ahk_id %id%	;Get the title of the window from its id
	; WinMinimize, %Title%
	; Return
; }

;===================
;=   SALAMANDER    =
;===================
;Numpad0::SwitchStartP("C:\Program Files\Altap Salamander\salamand.exe")

;===================
;=      SKYPE      =
;===================
;Numpad1::SwitchStartP("lync.exe")

;===================
;=     CHROME      =
;===================
;Numpad2::SwitchStartP("chrome.exe")

;===================
;=  TASK MANAGER   =
;===================
;Numpad3::SwitchStartP("taskmgr.exe")

;===================
;=   CALCULATOR    =
;===================
;Numpad4::SwitchStartP("calc.exe")

;launch calculator and paste clipboard
; ^Numpad4::
; {
	; SwitchStartPW("calc.exe", "Calculator", True)
	; Send, ^v
	; Return
; }

;===================
;=     HISTORY     =
;===================
Numpad5::KeyHistory

;===================
;=      HELP       =
;===================
;^Numpad5::DisplayHelp()

;===================
;=       AEC       =
;===================
;Numpad6::Run, "C:\Program Files\internet explorer\iexplore.exe" "http://msdn.microsoft.com/en-us/library/aa394372.aspx"

;Launches Notepad
;Numpad7::SwitchStartP("notepad.exe")

;===================
;=     OUTLOOK     =
;===================
;Activate Outlook and switch to mail
; Numpad8::
; {
	; SwitchStartP("outlook.exe", True)
	; Send, ^&
	; Return
; }

;Activate Outlook and search from clipboard
; ^Numpad8::
; {
	; Send, ^c
	; SwitchStartP("outlook.exe", True)
	; Send, ^e^v{Enter}
	; Return
; }

;Activate Outlook and create new mail
; #Numpad8::
; {
	; SwitchStartP("outlook.exe", True)
	; Send, ^&
	; Sleep, 1000
	; Send, ^n
	; Return
; }

;===================
;=   CLEARQUEST    =
;===================
;Numpad9::Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "http://msdn.microsoft.com/en-us/library/aa394372.aspx"

;===================
;=    SNIPPING     =
;===================
; PrintScreen::
; {
	; SwitchStartP("SnippingTool.exe", True)
	; Send, !n
	; Return
; }

;===================
;=     LOCK PC     =
;===================
;Pause::SendMessage, 0x112, 0xF170, 2,, Program Manager

;===================
;=  CLOSE WINDOW   =
;===================
;ScrollLock::!F4

;Right windows key
;Middle button click to task view
;RWin::Send, #{Tab}

;===================
;=      MOUSE      =
;===================
;Control + middle mouse button: switch mouse sensitivity
;MButton::SwitchMouseSensitivity()

;===================
;=    CapsLock     =
;===================
;=      DITTO      =
;^!v							;ditto paste text
;#<								;ditto paste position 2

; Appskey::Send, {Blind}^#ù		;Ditto
; CapsLock & f::					;Google search
; {
	; Send, {Blind}^c
	; Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://www.google.com/search?q="%clipboard%"
	; Return
; }

;= Window dragging =
; CapsLock & LButton::
; {
	; CoordMode, Mouse  ; Switch to screen/absolute coordinates.
	; MouseGetPos, EWD_MouseStartX, EWD_MouseStartY, EWD_MouseWin
	; WinGetPos, EWD_OriginalPosX, EWD_OriginalPosY,,, ahk_id %EWD_MouseWin%
	; WinGet, EWD_WinState, MinMax, ahk_id %EWD_MouseWin% 
	; if EWD_WinState = 0  ; Only if the window isn't maximized
		; SetTimer, EWD_WatchMouse, 10 ; Track the mouse as the user drags it.
	; Return

	; EWD_WatchMouse:
	; GetKeyState, EWD_LButtonState, LButton, P
	; if EWD_LButtonState = U  ; Button has been released, so drag is complete.
	; {
		; SetTimer, EWD_WatchMouse, Off
		; Return
	; }
	; GetKeyState, EWD_EscapeState, Escape, P
	; if EWD_EscapeState = D  ; Escape has been pressed, so drag is cancelled.
	; {
		; SetTimer, EWD_WatchMouse, Off
		; WinMove, ahk_id %EWD_MouseWin%,, %EWD_OriginalPosX%, %EWD_OriginalPosY%
		; Return
	; }
	;Otherwise, reposition the window to match the change in mouse coordinates
	;caused by the user having dragged the mouse:
	; CoordMode, Mouse
	; MouseGetPos, EWD_MouseX, EWD_MouseY
	; WinGetPos, EWD_WinX, EWD_WinY,,, ahk_id %EWD_MouseWin%
	; SetWinDelay, -1   ; Makes the below move faster/smoother.
	; WinMove, ahk_id %EWD_MouseWin%,, EWD_WinX + EWD_MouseX - EWD_MouseStartX, EWD_WinY + EWD_MouseY - EWD_MouseStartY
	; EWD_MouseStartX := EWD_MouseX  ; Update for the next timer-call to this subroutine.
	; EWD_MouseStartY := EWD_MouseY
	; Return
; }
;///////////////////
;/    CapsLock     /
;///////////////////

SwitchStartP(sProgramToStart, bWait:=False)
{
	sSearchWindow := "ahk_exe " . sProgramToStart
	SwitchStartPW(sProgramToStart, sSearchWindow, bWait)
}

SwitchStartPW(sProgramToStart, sSearchWindow, bWait:=False)
{
	if WinExist(sSearchWindow)
	{
		WinActivate, %sSearchWindow%
		;Traytip, Window, Switched to %sProgramToStart%, 0.5, 1
	}
	else
	{
		Run, %sProgramToStart%
		;Traytip, Window, Launched %sProgramToStart%, 0.5, 1
	}
	
	if(bWait = True)
		WinWaitActive, %sSearchWindow%
}

SwitchMouseSensitivity()
{
	static iSensitivity := 20
	sSensitivity = low
	
	if(iSensitivity = 20)
		iSensitivity = 3
	else
	{
		iSensitivity = 20
		sSensitivity = high
	}

	DllCall("SystemParametersInfo", Int,113, Int,0, UInt,iSensitivity, Int,2)
	
	Traytip, Mouse, Sensitivity set to %sSensitivity%, 0.5, 1
}

DisplayHelp()
{
	static bHelp := False
	
	if(bHelp = False)
	{
		bHelp := True
		SplashImage, Lcuts.png, b1 fs18, Lcuts: cutting the work :)
	}
	else
	{
		bHelp := False
		SplashImage, Off
	}
}
