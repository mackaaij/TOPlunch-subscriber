;Required for Internet Explorer
#include <IE.au3> 
;Required for GUI
#include <GUIConstants.au3>
;Required for DateDiff
#include <Date.au3>

$programtitle="TOPlunch subscriber 1.1"

TraySetToolTip($programtitle)

$inifile=@UserProfileDir & "\TOPlunch subscriber.ini"
If IniRead($inifile,"TOPlunch","NeverAsk","0") == 1	Then Exit ; Exit if one does not want to use this program
If IniRead($inifile,"TOPlunch", "LastDecided","0") == _NowCalcDate() Then Exit ; Exit if question is already answered today

If (@HOUR==11 AND @MIN>30) Or (@HOUR>11) Then Exit ; Don't do anything after 11:30

; $url = "http://intranet.topdesk.com/index.php?req=subscribeme2" ; The TOPlunch Subscriber v1.0 "subscribe me" URL (will auto login)
$url = "http://intranet.topdesk.com/index.php"
$snoozeTime=15 ; Minutes to snooze
dim $begin ; Global variable
$snoozing=False; Global variable

GUICreate("Subscribe to TOPlunch?",250,100) ; will create a dialog box that when displayed is centered
$Button_YES = GUICtrlCreateButton ("&Yes",  10, 10, 110,40)
$Button_NO = GUICtrlCreateButton ("&No",  120, 10, 110,40)
$Button_SNOOZE = GUICtrlCreateButton ("&Snooze " & $snoozeTime & " minutes", 10, 50,110,40)
If (@HOUR==11 AND @MIN>15) Or (@HOUR>11) Then GUICtrlSetState($Button_SNOOZE,$GUI_DISABLE) ; Disable the snooze button after 10:15
$Button_NEVER = GUICtrlCreateButton ("Never ask me again",  120, 50,110,40)
GUISetState ()      ; Display the dialog box

; Run the GUI until the dialog is closed
While 1
    $msg = GUIGetMsg()
    Select
        Case $msg = $GUI_EVENT_CLOSE
            ExitLoop
        Case $msg = $Button_YES
			GUICtrlSetState($Button_YES,$GUI_DISABLE)
			GUICtrlSetState($Button_NO,$GUI_DISABLE)
			GUICtrlSetState($Button_SNOOZE,$GUI_DISABLE)
			GUICtrlSetState($Button_NEVER,$GUI_DISABLE)
			GUICtrlSetData ($Button_YES,"Subscribing....")
			IniWrite ($inifile,"TOPlunch", "LastDecided", _NowCalcDate()) ; Remember when 'No' was decided to prevent repeating the question today
			
            ; _IECreate($url,0,1,0,0) ; Open Intranet subscribtion page (don't attach to an existing window, show the browser, return immediatly and bring into focus)
			; Line above was TOPlunch Subscriber v1.0 but then the intranet changed and required a POST as below
			$oIE = _IECreate($url) ; Open Intranet subscribtion page (don't attach to an existing window, show the browser, wait for the page load to complete and DON'T bring into focus)
			$oForm = _IEFormGetObjByName ($oIE, "subscribemeform")
			$oQuery = _IEFormElementGetObjByName ($oForm, "req")
			_IEFormElementSetValue ($oQuery, "subscribeme2")
			_IEFormSubmit ($oForm)

			Exit
        Case $msg = $Button_NO
			IniWrite ($inifile,"TOPlunch", "LastDecided", _NowCalcDate()) ; Remember when 'No' was decided to prevent repeating the question today
            Exit ; No? Then quit.
		Case $msg = $Button_SNOOZE
			Snooze($snoozeTime)
		Case $msg = $Button_NEVER
			IniWrite ($inifile,"TOPlunch", "NeverAsk", "1")
            Exit
	EndSelect
	
	If $snoozing==True Then
		If (TimerDiff($begin) > $SnoozeTime*1000*60) Then EndSnooze()
	EndIf
Wend

Func Snooze($SnoozeTime)
	GUICtrlSetState($Button_SNOOZE,$GUI_DISABLE) ; Disable the snooze button to indicate it is pressed

	$snoozing=True
	$begin = TimerInit() ; Start timer
	
	; Calculate end time for snooze (only to display on the button)
	$wakeDate = _DateAdd( 'n',$SnoozeTime, _NowCalc()) ; Add snooze minutes to current time
	Dim $wakeDatePart
	Dim $wakeTimePart
	_DateTimeSplit($wakeDate,$wakeDatePart,$wakeTimePart)
	$wakeTimePart[2]=StringFormat("%.2d",$wakeTimePart[2])
	$sleepmessage="Sleeping until " & $wakeTimePart[1] & ":" & $wakeTimePart[2]
	; $S =StringFormat ( "$String = %s" & @CRLF & "$Float = %.2f" & @CRLF & "$Int = %.2d" ,$String, $Float, $Int )
	GUICtrlSetData ($Button_SNOOZE,$sleepmessage)
	TraySetToolTip($programtitle & " - " & $sleepmessage)

	GUISetState(@SW_MINIMIZE) ; Minimize the window
EndFunc

Func EndSnooze()
	$snoozing=False
	GUICtrlSetData ($Button_SNOOZE,"&Snooze " & $snoozeTime & " minutes")
	If (@HOUR==11 AND @MIN<15) Or (@HOUR<11) Then GUICtrlSetState($Button_SNOOZE,$GUI_ENABLE) ; Enable the snooze button before 10:15
	GUISetState(@SW_RESTORE)
	TraySetToolTip($programtitle)
EndFunc