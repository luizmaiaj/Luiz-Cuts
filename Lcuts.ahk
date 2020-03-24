#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook ; to detect pressed keys
#InstallMouseHook ; to detect clicks
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Tray menu
Menu, Tray, Icon, Lcuts.ico ;Load icon
Menu, Tray, Tip, Luiz cuts ;Change display name

;Set Lock keys permanently
SetNumlockState, AlwaysOn
;SetCapsLockState, AlwaysOff
SetScrollLockState, AlwaysOff

;oReminder := new Ant( 30000 ) ;time is in milliseconds
oReminder := new Ant( 6000 ) ;for testing

Return

MButton::#Tab

;=    SHORTCUTS    =

; searches selected text on the internet
Appskey & f::
{
	Send, ^c
	Clipwait
	Run, "C:\Program Files\Mozilla Firefox\Firefox.exe" https://duckduckgo.com/?q="%clipboard%"
	Return
}

; translates selected text to english using google translate
Appskey & t::
{
	Send, ^c
	Clipwait
	Run, "C:\Program Files\Mozilla Firefox\Firefox.exe" https://translate.google.com/#auto/en/"%clipboard%"
	return
}

;lock pc and switch off monitor
F12::SendMessage, 0x112, 0xF170, 2,, Program Manager

;=     CLASSES     =
class Ant
{
    __New( aWhen )
	{
        this.period := aWhen
		this.timer := ObjBindMethod(this, "Tick")
		
        timer := this.timer
        SetTimer % timer, % this.period
    }
	
    Tick()
	{
		timer := this.timer
		
		WinGetActiveTitle, vTitle
		FormatTime, vDateFileName,, yyyy_MM_dd
		FormatTime, vDateContent,, dd/MM/yy
		FormatTime, vTimeContent,, HH:mm:ss
		
		vPeriod := this.period
		vPeriod := vPeriod / 60000
		
		vApp := getApp(vTitle)
		
		vFilename = C:\Users\l.maia\Documents\%vDateFileName%_usage_log.csv
		
		FileAppend, 
		(
		%vTitle%;%vPeriod%;%vDateContent%;%vTimeContent%;%vApp%`n
		), %vFilename%
		;MsgBox % "You asked to be reminded of: " . Filename
    }
}

;=    FUNCTIONS    =
getApp(vTitle)
{
	if StrLen(vTitle) = 0
		Return "Screen Off;Free;Private"

	; List the key-value pairs of an object
	vApps := Object("Firefox","Firefox;Busy;Private","Microsoft Visual Studio","Visual Studio;Busy;Internal","Visual Basic for Applications","VBA;Busy;Internal","Outlook","Outlook;Busy;Important","calendar","Outlook;Busy;Important","meeting","Outlook;Busy;Important","message","Outlook;Busy;Important","Notepad","Notepad;Busy;Internal","Excel","Excel;Busy;Important","Word","Word;Busy;Important","PowerPoint","PowerPoint;Busy;Important","Salamander","Salamander;Busy;Internal","Lock Screen","Lock Screen;Free;Private","Adobe","Adobe PDF;Busy;Important","Skype","Skype;Busy;Important","Slack","Slack;Busy;Internal","Autohotkey","Autohotkey;Busy;Private","OneNote","OneNote;Busy;Important")


	; The above expression could be used directly in place of "colours" below:
	for k, v in vApps
	{
		if instr(vTitle, k)
		{
			Return v
		}
	}

	Return "Other;Busy;Internal"
}
