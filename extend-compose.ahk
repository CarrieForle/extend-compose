#Include "VD.ah2"

#Requires AutoHotkey >=2
#SingleInstance Force
#WinActivateForce
brightnessStep := 3
A_MaxHotkeysPerInterval := 150
A_HotkeyIntervaOl := 1000
SetWinDelay -1
SetControlDelay -1
SetMouseDelay -1
SetDefaultMouseSpeed 0
SetKeyDelay -1
SendMode "Event"
ProcessSetPriority "A"
FileEncoding "UTF-8" ; https://www.autohotkey.com/docs/v2/lib/File.htm#Encoding
capsLockPresses := 0

;@Ahk2Exe-SetMainIcon resource\main.ico
;@Ahk2Exe-AddResource resource\suspend.ico, 206

if FileExist("xarty_config.ini") == ""
{
	FileAppend "
	(
    [xarty-global]

    ; Upon the Compose sequence complete typing,
    ; the window for you to press Compose key
    ; for the ranslation to begin. (in millisecond)

    ; This property should be a positive integer,
    ; otherwise invalid.
    intervalAllowedForComposeValidation = 4800

    ; When you press and release a modifer key,
    ; the window for you to press extend key that
    ; will still acdtivate a extend layer. (in millisecond)

    ; This property should be a positive integer,
    ; otherwise invalid.
    intervalAllowedForExtendLayerActivation = 400

    ; When you toggle suspension via keyboard,
    ; a beep is produced where high pitch means on
    ; and low is off.

    ; This property only accepts "true" and "false",
    ; otherwise invalid.
    beepOnToggleScriptSuspension = true

    ; How much does the brightness value change
    ; per key press.

    ; This property should be an integer 
    ; between 1 and 100
    brightnessStep = 3
	)", "xarty_config.ini", "UTF-16"
}

if FileExist("compose.txt") == ""
{
	FileAppend "
	(
	; This file is used to create compose key pairs
	; For details, specification, and guide of modification, refer to https://github.com/CarrieForle/xarty/wiki/Xarty-with-AHK#composetxt
	
	=btw=By the way
	=name=CarrieForle
	=lol=(ﾟ∀。)
	)", "compose.txt"
}


beepHelper()
{
	if A_IsSuspended
		SoundBeep 1600, 75
	else
		SoundBeep 1100, 75
}

; https://www.autohotkey.com/docs/v2/Functions.htm#ByRef
changeBrightness(brightness, timeout := 1)
{
	if (brightness > 0 && brightness < 100)
    {
		for property in ComObjGet("winmgmts:\\.\root\WMI").ExecQuery("SELECT * FROM WmiMonitorBrightnessMethods")
			property.WmiSetBrightness(timeout, brightness)
    } else if (brightness > 100)
    {
 		brightness := 100
 	} else if (brightness < 0)
    {
 		brightness := 0
 	}
}

getCurrentBrightNess()
{
	for property in ComObjGet( "winmgmts:\\.\root\WMI" ).ExecQuery( "SELECT * FROM WmiMonitorBrightness" )
		currentBrightness := property.CurrentBrightness	

	return currentBrightness
}

try
{
	intervalAllowedForComposeValidation := Integer(IniRead("xarty_config.ini", "xarty-global", "intervalAllowedForComposeValidation")),
	intervalAllowedForExtendLayerActivation := Integer(IniRead("xarty_config.ini", "xarty-global", "intervalAllowedForExtendLayerActivation"))
	beepOnToggleScriptSuspension := IniRead("xarty_config.ini", "xarty-global", "beepOnToggleScriptSuspension")
	brightnessStep := Integer(IniRead("xarty_config.ini", "xarty-global", "brightnessStep"))
    if beepOnToggleScriptSuspension = "true"
		toggleSuspension := () => (
			beepHelper(),	
			Suspend(-1)
		)
	else if beepOnToggleScriptSuspension = "false"
		toggleSuspension := () => suspend(-1)
	else
		throw ValueError("Invalid beepOnToggleScriptSuspension value")
	if intervalAllowedForComposeValidation <= 0 || intervalAllowedForExtendLayerActivation <= 0
		throw ValueError("Invalid intervalAllowedForComposeValidation value")
    if brightnessStep < 1 || brightnessStep > 100
        throw ValueError("Invalid brightnessStep value")
}
catch Error as e
{
	MsgBox e.Message . " Invalid or missing properties found in xarty_config.ini.`nThe program will be terminated."
	ExitApp
}

timeSinceLastKey := -intervalAllowedForComposeValidation - 1,
timeSinceExtendRestart := -intervalAllowedForExtendLayerActivation - 1,
maximumComposeKeyLength := 20

runAsAdmin(ItemName, ItemPos, MyMenu)
{
    try
    {
        if A_IsCompiled
            Run '*RunAs "' A_ScriptFullPath '" /restart'
        else
            Run '*RunAs "' A_AhkPath '" /restart "' A_ScriptFullPath '"'
    }
    catch Error as e
    {
        MsgBox e.Message . "Couldn't run as admin."
    }
}

A_TrayMenu.Add("See Compose", (ItemName, ItemPos, MyMenu) => Run("compose.txt"),),
A_TrayMenu.Add("Config", (ItemName, ItemPos, MyMenu) => Run("xarty_config.ini"),),
A_TrayMenu.Add("Run as Admin", runAsAdmin)

if (A_IsAdmin)
{
    A_TrayMenu.Disable("Run as Admin")
}

wordList := Array(),
wordList.Length := wordList.Capacity := maximumComposeKeyLength

loop wordList.Length
	wordList[A_Index] := Map()

Loop Read "compose.txt"
{
	if A_LoopReadLine == ""
	    continue
	
	delimiterChar := SubStr(A_LoopReadLine, 1, 1)
	
	if delimiterChar == ";" ||
	    delimiterChar == "`n" ||
	    delimiterChar == A_Tab ||
	    delimiterChar == A_Space
	    continue
	
	keypair := StrSplit(A_LoopReadLine, delimiterChar, , 3)
	
	if keypair.Length < 3 || keypair[2] == "" ||keypair[3] == ""
	{
		if "No" == MsgBox(A_LoopReadLine " is not a valid compose keypair.`n`nClick `"Yes`" to cuntinue and ignore this keypair.`nClick `"No`" to terminate the script.", "Error in compose.txt", 4)
			ExitApp
	}
	else if StrLen(keypair[2]) > maximumComposeKeyLength
	{
		if "No" == MsgBox(A_LoopReadLine " is too long for a key (> 10).`n`nClick `"Yes`" to cuntinue and ignore this keypair.`nClick `"No`" to terminate the script.", "Error in compose.txt", 4)
			ExitApp
	}
	else
	{
		wordList[maximumComposeKeyLength - StrLen(keypair[2]) + 1].Set(keypair[2], keypair[3])
	}
}

VarSetStrCapacity &keypair, 0

if GetKeyState("CapsLock", "T")
	SetCapsLockState "AlwaysOn"
else
	SetCapsLockState "AlwaysOff"

#InputLevel 1
#SuspendExempt
#HotIf A_IsSuspended
CapsLock::
{
    SetCapsLockState GetKeyState("CapsLock", "T") ? "AlwaysOff" : "ALwaysOn"

    KeyWait("CapsLock")
}
#HotIf ; for CapsLock to be toggleable while suspending
RAlt::toggleSuspension
^+sc006::
{
    SendEvent("{vk97 Up}{vk98 Up}{vk99 Up}"),
	SoundBeep(3000, 50),
	SoundBeep(4500, 50),
	Reload()
}
^+sc029::ExitApp
#SuspendExempt false

*CapsLock::
{
    global
	if A_PriorHotkey == "*CapsLock up" 
		&& A_TickCount - timeSinceExtendPrestart <= intervalAllowedForExtendLayerActivation
		capsLockPresses++
	else
		capsLockPresses := 0

	Switch Mod(capsLockPresses, 3)
	{
		Case 0:
			SendEvent "{vk97 DownR}{vk98 Up}{vk99 Up}"
		Case 1:
			SendEvent "{vk97 Up}{vk98 DownR}{vk99 Up}"
		Case 2:
			SendEvent "{vk97 Up}{vk98 Up}{vk99 DownR}"
	}

	timeSinceExtendPrestart := A_TickCount,
	KeyWait("CapsLock")
}

*CapsLock up::{
    SendEvent "{vk97 Up}{vk98 Up}{vk99 Up}"
}

#InputLevel 0

HoldKey(key) {
	SendInput("{Blind}{" key " DownR}"),
	KeyWait(key, "L")
}

HoldKeyNormal(key) {
	SendInput("{Blind}{" key " Down}"),
	KeyWait(key, "L")
}

MoveMouse(key, dir)
{
    SendMode "Event"
    SetMouseDelay 10
    SetDefaultMouseSpeed 0
    scale := 7

    while (GetKeyState(key, "P"))
    {
        switch dir, false
        {
            case "up":
                Click 0, -scale, , 0, "Relative"
            case "down":
                Click 0, scale, , 0, "Relative"
            case "left":
                Click -scale, 0, , 0, "Relative"
            case "right":
                Click scale, 0, , 0, "Relative"
            case "up left":
                Click -scale, -scale, , 0, "Relative"
            case "up right":
                Click scale, -scale, , 0, "Relative"
            case "down left":
                Click -scale, scale, , 0, "Relative"
            case "down right":
                Click scale, scale, , 0, "Relative"
        }
    }
}

vk97 & Esc::CapsLock
vk97 & F1::HoldKey "Media_Play_Pause"
vk97 & F1 up::Send "{blind}{Media_Play_Pause Up}"
vk97 & F2::HoldKey "Media_Prev"
vk97 & F2 up::Send "{blind}{Media_Prev Up}"
vk97 & F3::HoldKey "Media_Next"
vk97 & F3 up::Send "{blind}{Media_Next Up}"
vk97 & F4::HoldKey "Media_Stop"
vk97 & F4 up::Send "{blind}{Media_Stop Up}"
vk97 & F5::HoldKey "Volume_Mute"
vk97 & F5 up::Send "{blind}{Volume_Mute Up}"
vk97 & F6::Volume_Down
vk97 & F7::Volume_Up
vk97 & F8::HoldKey "Launch_Media"
vk97 & F8 up::Send "{blind}{Launch_Media Up}"
vk97 & F9::HoldKey "Launch_App2"
vk97 & F9 up::Send "{blind}{Launch_App2 Up}"
vk97 & F11::changeBrightness(getCurrentBrightNess() - brightnessStep)
vk97 & F12::changeBrightness(getCurrentBrightNess() + brightnessStep)

vk97 & sc029::PrintScreen
vk97 & sc002::MoveMouse("sc002", "left")
vk97 & sc003::MoveMouse("sc003", "right")
vk97 & sc004::LButton
vk97 & sc005::RButton
vk97 & sc006::MButton
vk97 & sc007::F6
vk97 & sc008::F7
vk97 & sc009::F8
vk97 & sc00a::F9
vk97 & sc00b::F10
vk97 & sc00c::F11
vk97 & sc00d::F12

vk97 & sc010::Esc
vk97 & sc011::WheelUp
vk97 & sc012::Browser_Back
vk97 & sc013::Browser_Forward
vk97 & sc014::MoveMouse("sc014", "up")
vk97 & sc015::PgUp
vk97 & sc016::Home
vk97 & sc017::Up
vk97 & sc018::End
vk97 & sc019::Del
vk97 & sc01a::WheelLeft
vk97 & sc01b::WheelRight

vk97 & sc01e::Alt
vk97 & sc01f::WheelDown
vk97 & sc020::Shift
vk97 & sc021::Ctrl
vk97 & sc022::MoveMouse("sc022", "down")
vk97 & sc023::PgDn
vk97 & sc024::Left
vk97 & sc025::Down
vk97 & sc026::Right
vk97 & sc027::BackSpace
vk97 & sc028::HoldKey "AppsKey"
vk97 & sc028 up::Send "{blind}{AppsKey Up}"
vk97 & sc02b::SendInput "{Click " A_ScreenWidth / 2 " " A_ScreenHeight / 2 " 0}"

vk97 & sc02c::SendInput "+{Home}{Backspace}"
vk97 & sc02d::SendInput "+{End}{Backspace}"
vk97 & sc02e::XButton1
vk97 & sc02f::XButton2
vk97 & sc030::Ins

vk97 & sc031::F1
vk97 & sc032::F2
vk97 & sc033::F3
vk97 & sc034::F4
vk97 & sc035::F5

vk97 & Enter::^BackSpace
vk97 & Space::Enter

goToRelativeDesktopNumIfNotOneDesktop(num) {
    if vd.GetCount() > 1 {
        VD.goToRelativeDesktopNum(num)
    }
}

goToDesktopNumIfNotOneDesktop(num) {
    if vd.GetCount() > 1 {
        VD.goToRelativeDesktopNum(num)
    }
}

vk97 & WheelDown::goToRelativeDesktopNumIfNotOneDesktop(1)
vk97 & WheelUp::goToRelativeDesktopNumIfNotOneDesktop(-1)
vk97 & LButton::goToRelativeDesktopNumIfNotOneDesktop(-1)
vk97 & RButton::{
    current_desktop_num := VD.getCurrentDesktopNum()
    
    if current_desktop_num == VD.GetCount() and VD.getCount() <= 8
    {
        VD.createDesktop(true)
    }
    else
    {
        VD.goToRelativeDesktopNum(1)
    }
}
vk97 & XButton1::goToDesktopNumIfNotOneDesktop(VD.GetCount())
vk97 & XButton2::goToDesktopNumIfNotOneDesktop(1)
vk97 & Left::goToRelativeDesktopNumIfNotOneDesktop(-1)
vk97 & Right::{
    current_desktop_num := VD.getCurrentDesktopNum()
    
    if current_desktop_num == VD.GetCount() and VD.getCount() <= 8
    {
        VD.createDesktop(true)
    }
    else
    {
        VD.goToRelativeDesktopNum(1)
    }
}
vk97 & Up::goToRelativeDesktopNumIfNotOneDesktop(-1)
vk97 & Down::goToRelativeDesktopNumIfNotOneDesktop(-1)

vk98 & sc002::!
vk98 & sc003::£
vk98 & sc004::€
vk98 & sc005::$
vk98 & sc006::%
vk98 & sc007::^
vk98 & sc008::Numpad7
vk98 & sc009::Numpad8
vk98 & sc00a::Numpad9
vk98 & sc00b::NumpadMult
vk98 & sc00c::NumpadSub
vk98 & sc00d::=
vk98 & sc010::Home
vk98 & sc011::Up
vk98 & sc012::End
vk98 & sc013::Del
vk98 & sc014::Esc
vk98 & sc015::PgUp
vk98 & sc016::Numpad4
vk98 & sc017::Numpad5
vk98 & sc018::Numpad6
vk98 & sc019::NumpadAdd
vk98 & sc01a::(
vk98 & sc01b::)
vk98 & sc02b::,
vk98 & sc01e::Left
vk98 & sc01f::Down
vk98 & sc020::Right
vk98 & sc021::Backspace
vk98 & sc022::NumLock
vk98 & sc023::PgDn
vk98 & sc024::Numpad1
vk98 & sc025::Numpad2
vk98 & sc026::Numpad3
vk98 & sc027::NumpadEnter
vk98 & sc028::'
vk98 & sc02c::^z
vk98 & sc02d::^x
vk98 & sc02e::^c
vk98 & sc02f::^v
vk98 & sc030::LButton
vk98 & sc031:::
vk98 & sc032::Numpad0
vk98 & sc033::Numpad0
vk98 & sc034::NumpadDot
vk98 & sc035::NumpadDiv

vk99 & sc029::
vk99 & sc002::
vk99 & sc003::
vk99 & sc004::
vk99 & sc005::
vk99 & sc006::
vk99 & sc007::
vk99 & sc008::
vk99 & sc009::
vk99 & sc00a::
vk99 & sc00b::
vk99 & sc00c::
vk99 & sc00d::return
vk99 & sc010::[
vk99 & sc011::]
vk99 & sc012::~
vk99 & sc013::
vk99 & sc014::
vk99 & sc015::return
vk99 & sc016::'
vk99 & sc017::"
vk99 & sc018::\
vk99 & sc019::
vk99 & sc01a::
vk99 & sc01b::
vk99 & sc02b::return
vk99 & sc01e::(
vk99 & sc01f::)
vk99 & sc020::`
vk99 & sc021::
vk99 & sc022::
vk99 & sc023::return
vk99 & sc024::`{
vk99 & sc025::}
vk99 & sc026::%
vk99 & sc027::!
vk99 & sc028::return
vk99 & sc02c::&
vk99 & sc02d::|
vk99 & sc02e::*
vk99 & sc02f::
vk99 & sc030::
vk99 & sc031::return
vk99 & sc032::+
vk99 & sc033::-
vk99 & sc034::=
vk99 & sc035::return

goToDesktopNumCheck(num) {
    if num <= VD.getCount() {
        VD.goToDesktopNum(num)
    }
}

#Numpad1::goToDesktopNumCheck(1)
#Numpad2::goToDesktopNumCheck(2)
#Numpad3::goToDesktopNumCheck(3)
#Numpad4::goToDesktopNumCheck(4)
#Numpad5::goToDesktopNumCheck(5)
#Numpad6::goToDesktopNumCheck(6)
#Numpad7::goToDesktopNumCheck(7)
#Numpad8::goToDesktopNumCheck(8)
#Numpad9::goToDesktopNumCheck(9)

ih := InputHook("V L" . maximumComposeKeyLength, "{Left}{Up}{Right}{Down}{Home}{PgUp}{End}{PgDn}"),
oldBuffer := ""

~RCtrl::
{	
	if A_PriorHotKey !== ThisHotKey
		onKeyDown(ih, 0x5d, 0x15d)
	KeyWait "RControl"
}

~Backspace::
~+Backspace::
{
	global oldBuffer
	if ih.Input == "" && oldBuffer != ""
		oldBuffer := SubStr(oldBuffer, 1, StrLen(oldBuffer) - 1)
}
~!Backspace::
~*^Backspace::
{
	if A_PriorHotKey != "~!Backspace" && A_PriorHotKey != "~*^Backspace"
	{
		ih.Stop(),
		ih.Start()
	}
}

onChar(ih, ch)
{
	global timeSinceLastKey := A_TickCount
}

onKeyDown(ih, vk, sc)
{
	global timeSinceLastKey
	
	if A_TickCount - timeSinceLastKey > intervalAllowedForComposeValidation
	{
		ih.Stop(),
		ih.Start()
	}
	
	else if A_TickCount - timeSinceLastKey <= intervalAllowedForComposeValidation
	{
		inpBuffer := oldBuffer . ih.Input
		for words in wordList
		{
			for key, val in words
			{
				if key == SubStr(inpBuffer, -StrLen(key))
				{
					ih.Stop(),
					SendInput("{Backspace " StrLen(key) "}" val),
					ih.Start()
					return
				}
			}
		}
	}
}

onEnd(ih)
{
	global oldBuffer
	if ih.EndReason == "Max"
	{
		oldBuffer := ih.Input
		ih.Start()
	}
	else
	{
		if ih.EndReason == "EndKey"
			ih.Start()
		oldBuffer := ""
	}
}

ih.OnKeyDown := onKeyDown,
ih.OnEnd := onEnd,
ih.OnChar := onChar,
ih.Start()