#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon_lunch.ico
#AutoIt3Wrapper_Outfile=TOPlunch_subscriber.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;Required for Internet Explorer
#include <IE.au3>
;Required for GUI
#include <GUIConstants.au3>
;Required for DateDiff
#include <Date.au3>
#include <Inet.au3>
#include <Array.au3>

$programtitle="TOPlunch subscriber 2.1"

TraySetToolTip($programtitle)

$inifile=@UserProfileDir & "\TOPlunch subscriber.ini"
if FileExists($inifile) Then
	If IniRead($inifile,"TOPlunch","NeverAsk","0") == 1	Then Exit ; Exit if one does not want to use this program
	If IniRead($inifile,"TOPlunch", "LastDecided","0") == _NowCalcDate() Then Exit ; Exit if question is already answered today
EndIf

If @HOUR=11 AND @MIN>15 Then exit; Don't do anything after 11:15
If @HOUR > 11 then exit

if StringInStr(@IPAddress1,"10.0.") <> 1 and StringInStr(@IPAddress2,"10.0.") <> 1 and StringInStr(@IPAddress3,"10.0.") <> 1 and stringInStr(@IPAddress4,"10.0.") <> 1 then Exit; do nothing when not in delft office

$url = "https://intranet-new.topdesk.com/subscribe.me.php" ; URL for Bitrix24
$snoozeTime=15 ; Minutes to snooze
dim $begin ; Global variable
$snoozing=False; Global variable

GUICreate("Subscribe to TOPlunch?",250,100) ; will create a dialog box that when displayed is centered
$Button_YES = GUICtrlCreateButton ("&Yes",  10, 10, 110,40)
$Button_NO = GUICtrlCreateButton ("&No",  120, 10, 110,40)
$Button_SNOOZE = GUICtrlCreateButton ("&Snooze " & $snoozeTime & " minutes", 10, 50,110,40)
If (@HOUR>11) Then GUICtrlSetState($Button_SNOOZE,$GUI_DISABLE) ; Disable the snooze button after 11:15
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

			$IECreate = _IECreate($url,0,1,0,0) ; Open Intranet subscribtion page (don't attach to an existing window, do show the browser, don't wait for the page to load, don't take focus)
            #comments-start
			$Output = _INetGetSource($url2)
			$FullName = GetFullName(@UserName)
			$SPlittedName = StringSplit($FullName," ")

		 if  StringInStr($Output,"Deze persoon luncht al mee!") = 0 then
			if StringInStr($Output,$SPlittedName[1],0,2) = 0 or StringInStr($Output,$SPlittedName[2],0,2) = 0  then
			   MsgBox(0,"Error during subscription","Not possible to subscribe!" & @LF & "Please check your settings in Internet Explorer" & @LF & "This can be done by pressing ALT X ->" & @LF & "Internet Options -> Connections -> LAN Settings" & @LF & '"Automatically detect settings" should be ticked')
			Else
            #comments-end
			   IniWrite ($inifile,"TOPlunch", "LastDecided", _NowCalcDate()) ; Remember when 'Yes' was decided to prevent repeating the question today
            #comments-start
			EndIf
		 EndIf
         #comments-end

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

Func GetFullName($sUserName)
    $colItems = ""
    $strComputer = "localhost"

    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount WHERE Name = '" & $sUserName &  "'", "WQL", 0x10 + 0x20)

    If IsObj($colItems) then
       For $objItem In $colItems
          Return $objItem.FullName
       Next
    Else
       Return SetError(1,0,"")
    Endif
EndFunc