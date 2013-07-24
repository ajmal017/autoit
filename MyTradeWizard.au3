#RequireAdmin
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator

#include <File.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <GuiMenu.au3>
#include <WindowsConstants.au3>

Opt("MustDeclareVars", 1)
Opt("ExpandVarStrings", 1)
Opt("GUIOnEventMode", 1)
Opt("WinTitleMatchMode", 2)
Opt("SendCapslockMode", 0)
Opt("SendKeyDelay", 50)
Opt("SendKeyDownDelay", 50)

Global Const $ONE_MINUTE = 60 * 1000

Global Const $MTW_TITLE = "My Trade Wizard"
Global Const $MTW = "[TITLE:$MTW_TITLE$; CLASS:AutoIt v3 GUI]"

Global Const $MTW_INI = ".\mtw.ini"
Global Const $MTW_LOG = ".\mtw.log"

Global Const $DO_YOU_WANT_TO_RUN_THIS_APPLICATION = "[TITLE:Security Information; CLASS:SunAwtDialog]"
Global Const $LOGIN_FRAME = "[TITLE:Login; CLASS:SunAwtFrame]"
Global Const $EXISTING_SESSION_DETECTED = "[TITLE:Existing session detected; CLASS:SunAwtDialog]"

Global Const $TWS_JNLP = ".\tws.jnlp"

Func RestartTWS()
   SetStatus("Restarting TWS...")
   StopTWS()
   If $fInterrupt Then
	  SetStatus("Interrupting TWS restart...")
	  Return
   EndIf
   StartTWS()
EndFunc

Func StopTWS()
   If WinExists($tws) Then
	  SetStatus("Closing TWS main window...")
	  While WinExists($tws)
		 WinKill($tws)
		 ConfirmExit()
	  WEnd
	  _FileWriteLog($MTW_LOG, "Stopped TWS")
   EndIf
   CleanUp()
   SetStatus("Disabled")
   $enabled = False
EndFunc

Func ConfirmExit()
   Local $tws_dialog = "[TITLE:$tws_title$; CLASS:SunAwtDialog]"
   If WinExists($tws_dialog) Then
	  SetStatus("Confirming TWS exit...")
	  ControlSend($tws_dialog, "", "", "{SPACE}")
	  ControlSend($tws_dialog, "", "", "{ESCAPE}")
   EndIf
EndFunc

Func CleanUp()
   While WinExists($DO_YOU_WANT_TO_RUN_THIS_APPLICATION)
	  SetStatus("Closing TWS confirm run dialog...")
	  WinKill($DO_YOU_WANT_TO_RUN_THIS_APPLICATION)
   WEnd
   
   While WinExists($LOGIN_FRAME)
	  SetStatus("Closing TWS login frame...")
	  WinKill($LOGIN_FRAME)
   WEnd
	  
   Local $title[11]
   $title[0]="Authenticating..."
   $title[1]="Requesting startup parameters..."
   $title[2]="Loading resources..."
   $title[3]="Loading..."
   $title[4]="Initializing environment..."
   $title[5]="Reading market rules..."
   $title[6]="Reading settings file..."
   $title[7]="Loading window factories..."
   $title[8]="Processing startup parameters..."
   $title[9]="Initializing managers..."
   $title[10]="Starting application..."
   For $i = 0 to UBound($title) - 1
	  Local $win = "[TITLE:" & $title[$i] & "; CLASS:SunAwtFrame]"
	  While WinExists($win)
		 SetStatus("Closing TWS login progress frame...")
		 WinKill($win)
	  WEnd
   Next
EndFunc

Func StartTWS()
   SetStatus("Starting TWS...")
   GUICtrlSetState($start_button, $GUI_DISABLE)
   GUICtrlSetState($stop_button, $GUI_ENABLE)
   While Not WinExists($tws)
	  CleanUp()
	  RunTwsJnlp()
	  Local $iBegin = TimerInit()
	  Do
		 If $fInterrupt Then
			SetStatus("Interrupting Login...")
			CleanUp()
			Return
		 ElseIf WinExists($LOGIN_FRAME) Then
			ExitLoop
		 Else
			SetStatus("Waiting for Login prompt...")
			BriefPause()
		 EndIf
	  Until TimerDiff($iBegin) > $ONE_MINUTE
	  Login()
	  Local $iBegin = TimerInit()
	  Do
		 If $fInterrupt Then
			SetStatus("Interrupting TWS start up...")
			CleanUp()
			Return
		 ElseIf WinExists($tws) Then
			SetStatus("Enabled")
			$enabled = True
			_FileWriteLog($MTW_LOG, "Started TWS")
			Return
		 ElseIf WinExists($EXISTING_SESSION_DETECTED) Then
			SetStatus("Disconnecting other session...")
			ControlSend($EXISTING_SESSION_DETECTED, "", "", "{SPACE}")
		 Else
			SetStatus("Waiting for TWS to start up...")
			BriefPause()
		 EndIf
	  Until TimerDiff($iBegin) > 5 * $ONE_MINUTE
   WEnd
EndFunc

Func RunTwsJnlp()
   DownloadTwsJnlp()
   If $fInterrupt Then
	  SetStatus("Interrupting TWS JNLP launch...")
	  Return
   EndIf
   SetStatus("Launching TWS JNLP...")
   ShellExecute($TWS_JNLP)
   ConfirmRun()
EndFunc

Func DownloadTwsJnlp()
   SetStatus("Downloading TWS JNLP...")
   Local Const $URL = "http://www.interactivebrokers.com/java/classes/tws.jnlp"
   InetGet($URL, $TWS_JNLP)
EndFunc

Func ConfirmRun()
   Local $iBegin = TimerInit()
   Do
	  If $fInterrupt Then
		 SetStatus("Interrupting Login...")
		 Return
	  ElseIf WinExists($DO_YOU_WANT_TO_RUN_THIS_APPLICATION) Or WinExists($LOGIN_FRAME) Then
		 ExitLoop
	  Else
		 BriefPause()
	  EndIf
   Until TimerDiff($iBegin) > $ONE_MINUTE
   If WinExists($DO_YOU_WANT_TO_RUN_THIS_APPLICATION) Then
	  SetStatus("Confirming TWS launch...")
	  While WinExists($DO_YOU_WANT_TO_RUN_THIS_APPLICATION)
		 ControlSend($DO_YOU_WANT_TO_RUN_THIS_APPLICATION, "", "", "{SPACE}")
	  WEnd
   EndIf
EndFunc

Func Login()
   If WinExists($LOGIN_FRAME) Then
	  SetStatus("Logging in...")
	  BlockInput(1)
	  ControlFocus($LOGIN_FRAME, "", "")
	  ControlSend($LOGIN_FRAME, "", "", "{CAPSLOCK off}")
	  ControlSend($LOGIN_FRAME, "", "", $login)
	  ControlSend($LOGIN_FRAME, "", "", "{TAB}")
	  ControlSend($LOGIN_FRAME, "", "", $password)
	  ControlSend($LOGIN_FRAME, "", "", "{ALTDOWN}o{ALTUP}")
	  BlockInput(0)
   EndIf
EndFunc

Func BriefPause()
   Sleep(250)
EndFunc

Func SetStatus($status)
   Local $status_bar_text = "Status: $status$"
   If Not (_GUICtrlStatusBar_GetText($hStatus, 0) = $status_bar_text) Then
	  _GUICtrlStatusBar_SetText($hStatus, $status_bar_text)
   EndIf
EndFunc

Func CloseOtherInstances()
   While WinExists($MTW)
	  WinClose($MTW)
	  WinWaitCLose($MTW)
   WEnd
EndFunc

Func IncreaseProcessPriority()
   Local Const $HIGH_PRIORITY = 4
   Local Const $PROCESS_NAME_WHEN_COMPILED = "MyTradeWizard.exe"
   Local Const $PROCESS_NAME_WHEN_SCRIPTED = "AutoIt3.exe"
   If ProcessExists($PROCESS_NAME_WHEN_COMPILED) Then
	  ProcessSetPriority($PROCESS_NAME_WHEN_COMPILED, $HIGH_PRIORITY)
   ElseIf ProcessExists($PROCESS_NAME_WHEN_SCRIPTED) Then
	  ProcessSetPriority($PROCESS_NAME_WHEN_SCRIPTED, $HIGH_PRIORITY)
   EndIf
EndFunc

Func CreateGui()
   Global $hGUI = GUICreate($MTW_TITLE, 325, 200)
   GUISetOnEvent($GUI_EVENT_CLOSE, "ClickExit")
   WinSetOnTop($MTW, "", 1)

   GUICtrlCreateGroup("Mode", 5, 5, 310, 40)
   Global $demo_radio = GUICtrlCreateRadio("Demo", 30, 20, 60, 20)
   GUICtrlSetOnEvent($demo_radio, "ClickDemo")
   Global $live_radio = GUICtrlCreateRadio("Live", 160, 20, 60, 20)
   GUICtrlSetOnEvent($live_radio, "ClickLive")

   GuiCtrlCreateGroup("Credentials", 5, 50, 310, 40)
   GUICtrlCreateLabel("Login:", 30, 65, 30, 20)
   Global $login = IniRead($MTW_INI, "Credentials", "Login", "")
   Global $login_input = GUICtrlCreateInput($login, 65, 65, 80, 20)
   GUICtrlCreateLabel("Password:", 160, 65, 60, 20)
   Global $password = IniRead($MTW_INI, "Credentials", "Password", "")
   Local Const $ES_PASSWORD = 0x0020
   Global $password_input = GUICtrlCreateInput($password, 215, 65, 80, 20, $ES_PASSWORD)
   
   GUICtrlCreateGroup("Restart At", 5, 95, 310, 40)
   Global $hour = IniRead($MTW_INI, "RestartAt", "Hour", "00")
   Global $hour_input = GUICtrlCreateInput($hour, 30, 110, 20, 20)
   GUICtrlCreateLabel(":", 54, 112, 5, 20)
   Global $minute = IniRead($MTW_INI, "RestartAt", "Minute", "00")
   Global $minute_input = GUICtrlCreateInput($minute, 60, 110, 20, 20)
   GUICtrlCreateLabel(":", 84, 112, 5, 20)
   Global $second = IniRead($MTW_INI, "RestartAt", "Second", "00")
   Global $second_input = GUICtrlCreateInput($second, 90, 110, 20, 20)

   Global $start_button = GUICtrlCreateButton("Start", 85, 145, 60, 20)
   GUICtrlSetState($start_button, $GUI_ENABLE)
   GUICtrlSetOnEvent($start_button, "ClickStart")
   Global $stop_button = GUICtrlCreateButton("Stop", 155, 145, 60, 20)
   GUICtrlSetState($stop_button, $GUI_DISABLE)
   GUICtrlSetOnEvent($stop_button, "ClickStop")
   
   Global $hStatus = _GUICtrlStatusBar_Create($hGUI)
   SetStatus("Disabled")

   GUISetState(@SW_SHOW)
   
   GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")
   GUIRegisterMsg($WM_SYSCOMMAND, "_WM_SYSCOMMAND")
   
   Local $mode = IniRead($MTW_INI, "General", "Mode", "Demo")
   If $mode = "Demo" Then
	  ClickDemo()
   ElseIf $mode = "Live" Then
	  ClickLive()
   EndIf
EndFunc

Func _WM_COMMAND($hWnd, $Msg, $wParam, $lParam)
   If BitAND($wParam, 0x0000FFFF) = $stop_button Then $fInterrupt = True
   Return $GUI_RUNDEFMSG
EndFunc

Func _WM_SYSCOMMAND($hWnd, $Msg, $wParam, $lParam)
   If BitAND($wParam, 0x0000FFFF) = $SC_CLOSE Then $fInterrupt = True
   Return $GUI_RUNDEFMSG
EndFunc

Func ClickExit()
   SetStatus("Exiting...")
   ClickStop()
   Exit
EndFunc

Func ClickDemo()
   GUICtrlSetState($demo_radio, $GUI_CHECKED)
   GUICtrlSetState($live_radio, $GUI_UNCHECKED)
   GUICtrlSetState($login_input, $GUI_DISABLE)
   GUICtrlSetState($password_input, $GUI_DISABLE)
   Global $tws_title = "IB TWS (Demo System)"
   SetTwsTitle()
EndFunc

Func ClickLive()
   GUICtrlSetState($demo_radio, $GUI_UNCHECKED)
   GUICtrlSetState($live_radio, $GUI_CHECKED)
   GUICtrlSetState($login_input, $GUI_ENABLE)
   GUICtrlSetState($password_input, $GUI_ENABLE)
   Global $tws_title = "IB Trader Workstation"
   SetTwsTitle()
EndFunc

Func SetTwsTitle()
   Global $tws = "[TITLE:$tws_title$; CLASS:SunAwtFrame]"
EndFunc

Func ClickStart()
   SetStatus("Handling Start Button")
   $fInterrupt = False
   If BitAND(GUICtrlRead($demo_radio), $GUI_CHECKED) = $GUI_CHECKED Then
	  Local $mode = "Demo"
	  $login = "edemo"
	  $password = "demouser"
   ElseIf BitAND(GUICtrlRead($live_radio), $GUI_CHECKED) = $GUI_CHECKED Then
	  Local $mode = "Live"
	  $login = GUICtrlRead($login_input)
	  $password = GUICtrlRead($password_input)
   EndIf
   $hour = GUICtrlRead($hour_input)
   $minute = GUICtrlRead($minute_input)
   $second = GUICtrlRead($second_input)
   IniWrite($MTW_INI, "General", "Mode", $mode)
   IniWrite($MTW_INI, "Credentials", "Login", GUICtrlRead($login_input))
   IniWrite($MTW_INI, "Credentials", "Password", GUICtrlRead($password_input))
   IniWrite($MTW_INI, "RestartAt", "Hour", $hour)
   IniWrite($MTW_INI, "RestartAt", "Minute", $minute)
   IniWrite($MTW_INI, "RestartAt", "Second", $second)
   RestartTWS()
EndFunc

Func ClickStop()
   SetStatus("Handling Stop Button")
   GUICtrlSetState($stop_button, $GUI_DISABLE)
   GUICtrlSetState($start_button, $GUI_DISABLE)
   StopTWS()
   GUICtrlSetState($start_button, $GUI_ENABLE)
EndFunc

CloseOtherInstances()
IncreaseProcessPriority()
CreateGui()

Global $fInterrupt = False
Global $enabled = False
While True
   If $enabled Then
	  SetStatus("Waiting until $hour$:$minute$:$second$ to restart TWS...")
	  If @HOUR = $hour And @MIN = $minute And @SEC = $second Then RestartTWS()
   Else
	  StopTWS()
   EndIf
   BriefPause()
WEnd