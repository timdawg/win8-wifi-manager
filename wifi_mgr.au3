#AutoIt3Wrapper_UseX64=N
#AutoIt3Wrapper_UseUpx=N
#AutoIt3Wrapper_Compression=4
;#AutoIt3Wrapper_Icon=.ico
;#AutoIt3Wrapper_OutFile=.exe
#AutoIt3Wrapper_Res_Description=Wireless Profile Manager
#AutoIt3Wrapper_Res_Language=1033

#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <Constants.au3>
#include <StringConstants.au3>
#include <ComboConstants.au3>
#include <GuiListView.au3>


#NoTrayIcon

Opt("TrayAutoPause",0)
AutoItSetOption("WinTitleMatchMode", 2)     ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
Opt("GUIOnEventMode", 1)  ; Change to OnEvent mode

Const $script_name = 'Wireless Profile Manager'

; Options
	$show_keys=False
	$sleep_delay=100
	$ini_path = StringReplace(StringReplace(@ScriptFullPath, ".exe", ".ini"), ".au3", ".ini")
; End Options

; Init Variables
	Global $gui_closed = False
	Global $gui_hnd = -1
	Global $aInterfaces[1]
	Global $aInterfacesHW[1]
	$aInterfaces[0]=0
	$aInterfacesHW[0]=0
	Global $aProfiles[1]
	Global $aProfilesSecurity[1]
	Global $aProfilesAuto[1]
	Global $aProfilesHidden[1]
	Global $aProfilesKey[1]
	$aProfiles[0]=0
	$aProfilesSecurity[0]=0
	$aProfilesAuto[0]=0
	$aProfilesHidden[0]=0
	$aProfilesKey[0]=0
	Global $list_items[1]
	$list_items[0]=0
; End Init Variables

; Create a GUI
$line_height = 15
$ctrl_width = 400
$ctrl_height = 400
;$ctrl_height = 170
$ctrl_left = (@DesktopWidth - $ctrl_width) / 2
$ctrl_top = (@DesktopHeight - $ctrl_height) / 2
$gui_hnd = GUICreate($script_name, $ctrl_width, $ctrl_height, $ctrl_left, $ctrl_top, BitOr($WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
GUISetOnEvent($GUI_EVENT_CLOSE, "EventHandler", $gui_hnd)
GUISetState(@SW_SHOW, $gui_hnd)

$cbo_interface = GUICtrlCreateCombo("", 10, 10, 380, -1, BitOr($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL, $WS_VSCROLL))
GUICtrlSetOnEvent($cbo_interface, "EventHandler")

$listview = GUICtrlCreateListView("Profile|Auto|Hidden|Security|Key", 10, 40, 380, 320)

$btn_up = GUICtrlCreateButton("Move Up", 10, 370, 80, 20)
GUICtrlSetOnEvent($btn_up, "EventHandler")

$btn_down = GUICtrlCreateButton("Move Down", 100, 370, 80, 20)
GUICtrlSetOnEvent($btn_down, "EventHandler")

$btn_apply = GUICtrlCreateButton("Apply", 190, 370, 50, 20)
GUICtrlSetOnEvent($btn_apply, "EventHandler")

#cs
$btn_delete = GUICtrlCreateButton("Delete", 250, 350, 50, 20)
GUICtrlSetOnEvent($btn_delete, "EventHandler")

$btn_advanced = GUICtrlCreateButton("Advanced", 340, 350, 60, 20)
GUICtrlSetOnEvent($btn_advanced, "EventHandler")

$btn_add = GUICtrlCreateButton("Add", 430, 350, 50, 20)
GUICtrlSetOnEvent($btn_add, "EventHandler")
#ce

get_interfaces()
get_profiles()

While True
	If $gui_closed Then
		Exit
	EndIf

	Sleep($sleep_delay)

WEnd

Func EventHandler()
	Switch @GUI_CtrlId
		Case $GUI_EVENT_CLOSE
			$gui_closed = True
			GUISetState(@SW_HIDE, $gui_hnd)
		Case $cbo_interface
			get_profiles()
		Case $btn_up
			MoveUp()
		Case $btn_down
			MoveDown()
		Case $btn_apply
			apply_order()
		#cs
		Case $btn_disable
			DisableHost()
		Case $btn_delete
			DeleteHost()
		Case $btn_advanced
			EditIniFile()
		#ce
	EndSwitch
EndFunc

Func MoveUp()
	$sel_idx = _GUICtrlListView_GetSelectionMark($listview)
	$profile_idx = $sel_idx + 1
	$swap_idx = $profile_idx - 1

	If $profile_idx > 1 Then
		_ArraySwap($aProfiles, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesAuto, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesSecurity, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesHidden, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesKey, $profile_idx, $swap_idx)
	EndIf
	update_profiles()
	_GUICtrlListView_SetItemSelected($listview, ($swap_idx - 1), True, True)
EndFunc

Func MoveDown()
	$sel_idx = _GUICtrlListView_GetSelectionMark($listview)
	$profile_idx = $sel_idx + 1
	$swap_idx = $profile_idx + 1

	If $profile_idx < $aProfiles[0] And $profile_idx > 0 Then
		_ArraySwap($aProfiles, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesAuto, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesSecurity, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesHidden, $profile_idx, $swap_idx)
		_ArraySwap($aProfilesKey, $profile_idx, $swap_idx)
	EndIf
	update_profiles()
	_GUICtrlListView_SetItemSelected($listview, ($swap_idx - 1), True, True)
EndFunc

Func update_profiles()
	; Clear Old Values
	_GUICtrlListView_DeleteAllItems($listview)
	While $list_items[0] > 0
		_ArrayDelete($list_items, $list_items[0])
		$list_items[0] -= 1
	WEnd

	; Repopulate List View
	For $x = 1 To $aProfiles[0]
		$sListItem = $aProfiles[$x]
		$sListItem &= '|' & $aProfilesAuto[$x]
		$sListItem &= '|' & $aProfilesHidden[$x]
		$sListItem &= '|' & $aProfilesSecurity[$x]
		$sListItem &= '|' & $aProfilesKey[$x]
		_ArrayAdd($list_items, GUICtrlCreateListViewItem($sListItem, $listview))
		$list_items[0] += 1
	Next
	autosize_columns($listview)
EndFunc


Func autosize_columns($in_listview)
	$padding = 20
	$cnt = _GUICtrlListView_GetColumnCount

	For $iCol = 0 to $cnt - 1
		_GUICtrlListView_SetColumnWidth($in_listview, $iCol, $LVSCW_AUTOSIZE)
		$w1 = _GUICtrlListView_GetColumnWidth($in_listview, $iCol)
		_GUICtrlListView_SetColumnWidth($in_listview, $iCol, $LVSCW_AUTOSIZE_USEHEADER)
		$w2 = _GUICtrlListView_GetColumnWidth($in_listview, $iCol)
		If $w2 > $w1 Then
			_GUICtrlListView_SetColumnWidth($in_listview, $iCol, ($w2 + $padding))
		Else
			_GUICtrlListView_SetColumnWidth($in_listview, $iCol, ($w1 + $padding))
		EndIf
	Next
EndFunc


Func get_interfaces()
	$std_out = ''
	$pid = Run('netsh.exe wlan show interfaces', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)

	While 1
		$line = StdoutRead($pid)
		if @extended > 0 Then
			$std_out = $std_out & $line

		EndIf
		If @error Then ExitLoop
		Sleep(10)
	Wend

	If $std_out <> '' Then
		$aLines = StringSplit($std_out, @CRLF)

		For $x = 1 To $aLines[0]
			$line = $aLines[$x]
			$aLineData = StringSplit($line, ' : ', $STR_ENTIRESPLIT)
			If IsArray($aLineData) Then
				If $aLineData[0] > 1 Then
					$line_name = StringStripWS($aLineData[1], $STR_STRIPLEADING + $STR_STRIPTRAILING)
					$line_value = StringStripWS($aLineData[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)
					Switch $line_name
						Case 'Name'
							_ArrayAdd($aInterfaces, $line_value)
							$aInterfaces[0] += 1
						Case 'Description'
							_ArrayAdd($aInterfacesHW, $line_value)
							$aInterfacesHW[0] += 1
					EndSwitch
				EndIf
			EndIf
		Next
	EndIf

	$sDefault = ''
	$sCboData = ''
	For $x = 1 To $aInterfaces[0]
		If $x > 1 Then
			$sCboData &= '|'
		EndIf
		$sCboData &= $aInterfaces[$x] & '  [' & $aInterfacesHW[$x] & ']'
		If $x = 1 Then
			$sDefault = $sCboData
		EndIf
	Next
	GUICtrlSetData($cbo_interface, $sCboData, $sDefault)
EndFunc

Func clear_array(ByRef $in_array)
	While $in_array[0] > 0
		_ArrayDelete($in_array, $in_array[0])
		$in_array[0] -= 1
	WEnd
EndFunc

Func get_profiles()
	; Clear Arrays
	clear_array($aProfiles)
	clear_array($aProfilesSecurity)
	clear_array($aProfilesAuto)
	clear_array($aProfilesHidden)
	clear_array($aProfilesKey)

	; Get Interface Name (from Dropdown)
	$sInterface = StringRegExpReplace(GUICtrlRead($cbo_interface), '(.*)  \[.*\]', '$1')

	$std_out = ''
	$pid = Run('netsh.exe wlan show profiles interface="' & $sInterface & '"', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)

	While 1
		$line = StdoutRead($pid)
		if @extended > 0 Then
			$std_out = $std_out & $line

		EndIf
		If @error Then ExitLoop
		Sleep(10)
	Wend

	If $std_out <> '' Then
		$bUserProfiles = False
		$aLines = StringSplit($std_out, @CRLF)

		For $x = 1 To $aLines[0]
			$line = StringStripWS($aLines[$x], $STR_STRIPLEADING + $STR_STRIPTRAILING)
			If $bUserProfiles Then
				$aLineData = StringSplit($line, ' : ', $STR_ENTIRESPLIT)
				If IsArray($aLineData) Then
					If $aLineData[0] > 1 Then
						$profile_name = StringStripWS($aLineData[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)
						_ArrayAdd($aProfiles, $profile_name)
						$aProfiles[0] += 1
						get_profile_info($profile_name)
					EndIf
				EndIf
			ElseIf $line = 'User profiles' Then
				$bUserProfiles = True
			EndIf
		Next
	EndIf

	update_profiles()
EndFunc

Func get_profile_info($profile_name)
	; Create array item with blank value (in case of error)
	_ArrayAdd($aProfilesSecurity, '')
	$aProfilesSecurity[0] += 1
	_ArrayAdd($aProfilesAuto, '')
	$aProfilesAuto[0] += 1
	_ArrayAdd($aProfilesHidden, '')
	$aProfilesHidden[0] += 1
	_ArrayAdd($aProfilesKey, '')
	$aProfilesKey[0] += 1

	; Get Interface Name (from Dropdown)
	$sInterface = StringRegExpReplace(GUICtrlRead($cbo_interface), '(.*)  \[.*\]', '$1')

	; For more info about each profile use: netsh wlan show profiles name="NETWORK"
	; To include password add: key=clear
	$std_out = ''

	$netsh_params = 'wlan show profiles name="' & $profile_name & '" interface="' & $sInterface & '"'
	If $show_keys Then
		$netsh_params &= ' key=clear'
	EndIf
	$pid = Run('netsh.exe ' & $netsh_params, @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)

	While 1
		$line = StdoutRead($pid)
		if @extended > 0 Then
			$std_out = $std_out & $line

		EndIf
		If @error Then ExitLoop
		Sleep(10)
	Wend

	If $std_out <> '' Then
		$aLines = StringSplit($std_out, @CRLF)

		For $x = 1 To $aLines[0]
			$line = $aLines[$x]
			$aLineData = StringSplit($line, ' : ', $STR_ENTIRESPLIT)
			If IsArray($aLineData) Then
				If $aLineData[0] > 1 Then
					$line_name = StringStripWS($aLineData[1], $STR_STRIPLEADING + $STR_STRIPTRAILING)
					$line_value = StringStripWS($aLineData[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)
					Switch $line_name
						Case 'Authentication'
							$aProfilesSecurity[$aProfilesSecurity[0]] = $line_value
						Case 'Connection mode'
							If $line_value = "Connect automatically" Then
								$aProfilesAuto[$aProfilesAuto[0]] = 'Y'
							EndIf
						Case 'Network broadcast'
							If $line_value = "Connect even if this network is not broadcasting" Then
								$aProfilesHidden[$aProfilesHidden[0]] = 'Y'
							EndIf
						Case 'Security key'
							If $show_keys = False Then
								$aProfilesKey[$aProfilesKey[0]] = $line_value
							EndIf
						Case 'Key Content'
							If $show_keys Then
								$aProfilesKey[$aProfilesKey[0]] = $line_value
							EndIf
					EndSwitch
				EndIf
			EndIf
		Next
	EndIf
EndFunc


Func apply_order()
	; Init Error Flag
	$error = False

	; Get Interface Name (from Dropdown)
	$sInterface = StringRegExpReplace(GUICtrlRead($cbo_interface), '(.*)  \[.*\]', '$1')

	; Run netsh command for each profile
	For $x = 1 To $aProfiles[0]
		$profile_name = $aProfiles[$x]
		$netsh_params = 'wlan set profileorder name="' & $profile_name & '" interface="' & $sInterface & '" priority=' & $x

		$exitcode = RunWait('netsh.exe ' & $netsh_params, @SystemDir, @SW_HIDE)
		If $exitcode <> 0 Then
			$error = True
		EndIf
	Next

	If $error Then
		MsgBox($MB_OK + $MB_ICONERROR, 'Error - ' & $script_name, 'An error occurred while trying to apply the profile order!')
	Else
		MsgBox($MB_OK + $MB_ICONINFORMATION, $script_name, 'The profile order was successfully updated!')
	EndIf

	; Update profile list view
	get_profiles()
EndFunc

