#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook ; to detect pressed keys
#InstallMouseHook ; to detect clicks
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Tray menu
Menu, Tray, Icon, Lcuts.ico ;Load icon
Menu, Tray, Tip, Luiz cuts ;Change display name

;Contextual menu
Menu, MainMenu, Add, Help, MenuHandler
Menu, MainMenu, Icon, Help, Lcuts.ico ;Load icon
Menu, MainMenu, Add, Outlook, MenuHandler
Menu, MainMenu, Add, Slack, MenuHandler
Menu, MainMenu, Add, Salamander, MenuHandler
Menu, MainMenu, Add, Jira Board, MenuHandler
Menu, MainMenu, Add, Jira Filters, MenuHandler
Menu, MainMenu, Add, Confluence NR, MenuHandler
;Menu, MainMenu, Add, Reminder, MenuHandler

Menu, ReminderMenu, Add, 5 minutes, MenuHandler
Menu, ReminderMenu, Add, 30 minutes, MenuHandler
Menu, ReminderMenu, Add, 2 hours, MenuHandler

Menu, MainMenu, Add, Reminder, :ReminderMenu

;Set Lock keys permanently
SetNumlockState, AlwaysOn
SetCapsLockState, AlwaysOff
SetScrollLockState, AlwaysOff

Return

Numpad1::DisplayHelp()
Numpad2::SwitchStartP("chrome.exe")
Numpad3::SwitchStartP("excel.exe")
+Capslock::SwitchStartP("C:\Program Files\Altap Salamander\salamand.exe")

Capslock::SwitchOutlookSlack()

MButton::#Tab

PrintScreen::
{
	SwitchStartP("SnippingTool.exe", True)
	Send, !n
	Return
}

Appskey::WinMinimize, A

;Appskey::
;	Menu, MainMenu, Show
;Return

MenuHandler:
{
	if A_ThisMenuItem = Help
	{
		DisplayHelp()
	}
	else if A_ThisMenuItem = Outlook
	{
		SwitchStartP("outlook.exe", True)
	}
	else if A_ThisMenuItem = Slack
	{
		SwitchStartPW("C:\Users\l.maia\AppData\Local\slack\slack.exe", "slack.exe")
	}
	else if A_ThisMenuItem = Salamander
	{
		SwitchStartP("C:\Program Files\Altap Salamander\salamand.exe")
	}
	else if A_ThisMenuItem = Jira Board
	{
		Run, chrome.exe https://transurbsimulation.atlassian.net/secure/RapidBoard.jspa?rapidView=69&projectKey=LNM
	}
	else if A_ThisMenuItem = Jira Filters
	{
		Run, chrome.exe https://transurbsimulation.atlassian.net/secure/ManageFilters.jspa?owner=l.maia&page=1&selectedGroup=anything&selectedProject=anything&sortKey=name&sortOrder=ASC
	}
	else if A_ThisMenuItem = Confluence NR
	{
		Run, chrome.exe https://transurbsimulation.atlassian.net/wiki/spaces/L091NE/overview
	}
	else if A_ThisMenuItem = 5 minutes
	{
		InputBox, UserInput, Reminder, What do you want to be reminded of?, , , 100
			oReminder := new Elephant( 5 * 60000, UserInput )
	}
	else if A_ThisMenuItem = 30 minutes
	{
		InputBox, UserInput, Reminder, What do you want to be reminded of?, , , 100
			oReminder := new Elephant( 30 * 60000, UserInput )
	}
	else if A_ThisMenuItem = 2 hours
	{
		InputBox, UserInput, Reminder, What do you want to be reminded of?, , , 100
			oReminder := new Elephant( 2 * 60 * 60000, UserInput )
	}

	Return
}

;===================
;=     HISTORY     =
;===================
Numpad5::KeyHistory

;===================
;=    FUNCTIONS    =
;===================
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
		
	Return
}

SwitchOutlookSlack()
{
	static iComm := 1
	
	if(iComm = 1)
	{
		SwitchStartP("outlook.exe", True)
		iComm = 2
	}
	else
	{
		SwitchStartPW("C:\Users\l.maia\AppData\Local\slack\slack.exe", "slack.exe")
		iComm = 1
	}

	Return
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
	
	Return
}

DisplayHelp()
{
	Traytip, LCuts help, Capslock: Outlook and Slack`nShift+Capslock:Salamander`nNumPad2:Chrome`nNumPad3:Excel, 0.5, 1
}

class Elephant
{
    __New( aWhen, aWhat )
	{
        this.period := aWhen
        this.what := aWhat
        this.timer := ObjBindMethod(this, "Expire")
		
        timer := this.timer
        SetTimer % timer, % this.period
    }
    Expire()
	{
		timer := this.timer
        SetTimer % timer, Off
		MsgBox % "You asked to be reminded of: " . this.what
    }
}

;Counter class
class SecondCounter
{
    __New( aInterval )
	{
        this.interval := aInterval
        this.count := 0
        ; Tick() has an implicit parameter "this" which is a reference to the object, so we need to create a function which encapsulates "this" and the method to call:
        this.timer := ObjBindMethod(this, "Tick")
		
		; Known limitation: SetTimer requires a plain variable reference.
        timer := this.timer
        SetTimer % timer, % this.interval
        ToolTip % "Counter started"
    }
	__Delete()
	{
	}
    Stop()
	{
        ; To turn off the timer, we must pass the same object as before:
        timer := this.timer
        SetTimer % timer, Off
        ToolTip % "Counter stopped at " this.count
    }
    ; In this example, the timer calls this method:
    Tick()
	{
        ToolTip % ++this.count
    }
}