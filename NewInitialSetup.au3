#RequireAdmin
#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include<Array.au3>


Global $spath = @ScriptDir
Global $DOS, $Message, $dst, $bkpTime,  $logEnable = ""
$sKeyPath = "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook"
$cfgFile = @ScriptDir & "\Config.ini"
$log = False
Global $logEnable = "False"
Global $objArray[1][2] = [[0,0]]
FileChangeDir($spath)

Global $installdir = "C:\OutlookAuto\"
$Desktop = @DesktopDir

If FileExists(@scriptdir & "\Config.ini") Then
	  Global $logEnable =  IniRead($cfgFile, "Section", "LogEnable", "False")
EndIf

$dread = IniRead($cfgFile,'Section','DestinationFolder','D:\Outlook_Backup')
$timeread = IniRead($cfgFile,'Section','BackupTime','20:00:00')

Func SendAndLog($Data, $FileName = -1, $TimeStamp = True, $Log = $logEnable)
   If StringCompare($Log, 'True') ==0 Then

	   If $FileName == -1 Then $FileName = @ScriptDir & '\Log.txt'
	   ;Send($Data)
	   $hFile = FileOpen($FileName, 1)
	   If $hFile <> -1 Then
		   If $TimeStamp = True Then $Data = _Now() & ' - ' & $Data
		   FileWriteLine($hFile, $Data)
		   FileClose($hFile)
		EndIf
   EndIf
EndFunc


Func task_schedule_delete()
   SendAndLog("[task_schedule_delete]: Deleting task OutlookAuto")
   $DOS = Run(@ComSpec & ' /c schtasks /delete /tn OutlookAuto /f', "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
   ProcessWaitClose($DOS)
   SendAndLog("[task_schedule_delete]: reseting config")
EndFunc


task_schedule_delete()

#Region ### START Koda GUI section ### Form=
Global $InitialConfig_1 = GUICreate("Initial Config", 634, 480, -1, -1)
Global $Company = GUICtrlCreatePic("C:\FreeLance\img\soft.bmp", 0, 56, 76, 422)
Global $Outlook = GUICtrlCreatePic("C:\FreeLance\img\try.jpg", 0, 0, 628, 60)
Global $Blank = GUICtrlCreatePic("C:\FreeLance\img\black.bmp", -8, 56, 636, 4)
Global $Save = GUICtrlCreateButton("Save and Run", 480, 368, 91, 25)
Global $Cancel = GUICtrlCreateButton("Cancel", 390, 368, 75, 25)


;----------------------------------------------------------------------
; Inputs for Mail and destination Folders
;----------------------------------------------------------------------

Global $Mail_ID = GUICtrlCreateInput('', 296, 120, 121, 21)
Global $Mail_ID_Label = GUICtrlCreateLabel("Outlook Mail ID", 136, 128, 99, 17)
;GUICtrlSendMsg($Mail_ID, $EM_SETCUEBANNER, False, "adiprasoft@gmail.com") ; place holder
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $DestinationFolder = GUICtrlCreateInput('', 296, 176, 121, 21)
Global $DestinationLabel = GUICtrlCreateLabel("Destination Folder", 136, 184, 107, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $Select_Folder = GUICtrlCreateButton("...", 424, 176, 27, 21)
GUICtrlSetCursor (-1, 0)


;----------------------------------------------------------------------
; Inputs for Backup time Delete Mail Option
;----------------------------------------------------------------------
Global $BackupLabel = GUICtrlCreateLabel("Backup Time", 136, 240, 78, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $BackupupTime = GUICtrlCreateDate("20:00:00", 296, 232, 121, 21,$DTS_TIMEFORMAT)
GUICtrlSetCursor (-1, 0)
Global $Label1 = GUICtrlCreateLabel("Delete Mails After Backup", 136, 304, 152, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
Global $DeleteMailsOptions = GUICtrlCreateGroup("", 296, 288, 137, 41)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $Radio1 = GUICtrlCreateRadio("Yes", 304, 304, 49, 17)
GUICtrlSetCursor (-1, 0)
Global $Radio2 = GUICtrlCreateRadio("No", 368, 304, 41, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetCursor (-1, 0)



;----------------------------------------------------------------------
; Inputs for LogEnable Options
;----------------------------------------------------------------------
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $LogEnable = GUICtrlCreateCheckbox("Enable Logs", 136, 360, 97, 17)
Global $Silent = GUICtrlCreateCheckbox("Silent Mode", 136, 384, 97, 17)


;----------------------------------------------------------------------
; Inputs for Include Mail Options
;----------------------------------------------------------------------

Global $Group1 = GUICtrlCreateGroup("", 480, 96, 105, 162)
Global $Include_Label = GUICtrlCreateLabel("Also Include:", 496, 112, 81, 17)
GUICtrlSetFont(-1, 8, 800, 4, "MS Sans Serif")


Global $Inbox_Check = GUICtrlCreateCheckbox("Inbox", 496, 136, 81, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetCursor (-1, 7)

Global $Sent_Items_Include = GUICtrlCreateCheckbox("Sent Items", 496, 160, 73, 17)
GUICtrlSetCursor (-1, 0)
Global $Deleted_Items_include = GUICtrlCreateCheckbox("Deleted Items", 496, 184, 81, 17)
GUICtrlSetCursor (-1, 0)
Global $Drafts_Include = GUICtrlCreateCheckbox("Drafts", 496, 208, 49, 17)
GUICtrlSetCursor (-1, 0)
Global $Junk_Email_Include = GUICtrlCreateCheckbox("Junk Email", 496, 232, 73, 17)
GUICtrlSetCursor (-1, 0)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###




While 1
	$nMsg = GUIGetMsg()

	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Select_Folder
			_select_folder()

		Case $Cancel
			Exit

		Case $Save


			$mail_id = GUICtrlRead($Mail_ID)
			$dest_path = GUICtrlRead($DestinationFolder)

			$backupTimer = GUICtrlRead($BackupupTime)

			; writing to config file
			$mail_write = IniWrite("Config.ini","Section","MailID",$mail_id)
			$dwrite = IniWrite("Config.ini","Section",'DestinationFolder',$dest_path)
			$timewriting = IniWrite("Config.ini","Section",'BackupTime',$backupTimer)

			If GUICtrlRead($LogEnable) = $GUI_CHECKED Then
				$Logwrite = IniWrite("Config.ini","Section",'LogEnable','True')
			Else
				$Logwrite = IniWrite("Config.ini","Section",'LogEnable','False')
			EndIf

			If GUICtrlRead($Silent) = $GUI_CHECKED Then
				$Silentwrite = IniWrite("Config.ini","Section",'AutoSilent','True')
			Else
				$Silentwrite = IniWrite("Config.ini","Section",'AutoSilent','False')
			EndIf




			Global $Include_list[1][1] = [['Inbox']]

			If GUICtrlRead($Sent_Items_Include) == $GUI_CHECKED Then
				ArrayAdd('Sent Items')
			EndIf

			If GUICtrlRead($Deleted_Items_include) == $GUI_CHECKED Then
				ArrayAdd('Deleted Items')
			EndIf


			If GUICtrlRead($Drafts_Include) == $GUI_CHECKED Then
				ArrayAdd('Drafts')
			EndIf

			If GUICtrlRead($Junk_Email_Include) == $GUI_CHECKED Then
				ArrayAdd('Junk Email')
			EndIf


			$Outlook_Include_list = IniWrite($cfgFile,"Section","Include_List",_ArrayToString($Include_list,","))
			;_ArrayDisplay($Include_list)

			$createShortcut = FileCreateShortcut($installdir & "InitialSetup.exe",@DesktopDir & "\OutlookAuto.lnk","","","OutlookAuto",$installdir & "OutlookBackup.ico")
			FileChangeDir($spath)
			;$Enginestart = Run('Outlook_auto.exe')
		Exit
	EndSwitch
WEnd

;====================================================================
; FUNCs
;====================================================================

FUNC _select_folder()
    $sDir = FileSelectFolder("Select Destination Directory", "", 0, @ScriptDir)
    If StringRight($sDir, 1) <> "\" Then $sDir &= "\"

    If $sDir = "\" Then
        ToolTip("user abort",0,0)
        Sleep(500)
        ToolTip("",0,0)
    Elseif $sDir <> "\" Then
	    GUICtrlSetData($DestinationFolder,$sDir)
    EndIf
EndFunc




Func ArrayAdd($subfolder)
	;$Include_list[0][0] +=1
	_ArrayAdd($Include_list, $subfolder)
EndFunc








