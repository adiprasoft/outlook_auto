#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=img\OutlookBackup.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>

Global $spath = @ScriptDir
FileChangeDir($spath)

Global $installdir = "C:\OutlookAuto\"
$Desktop = @DesktopDir

$sread = IniRead('Config.ini','Section','SourceFolder','Auto')
$dread = IniRead('Config.ini','Section','DestinationFolder','D:\OutlookBackup')
$timeread = IniRead('Config.ini','Section','BackupTime','19:00:00')






#Region ### START Koda GUI section ### Form=
$InitialConfig = GUICreate("Initial Config", 631, 457, -1, -1)
$Company = GUICtrlCreatePic("C:\FreeLance\img\soft.bmp", -8, 56, 76, 398)
$Outlook = GUICtrlCreatePic("C:\FreeLance\img\try.jpg", 0, 0, 628, 60)
$Blank = GUICtrlCreatePic("C:\FreeLance\img\black.bmp", -8, 56, 636, 4)
$Save = GUICtrlCreateButton("Save and Run", 448, 344, 91, 25)
$Cancel = GUICtrlCreateButton("Cancel", 352, 344, 75, 25)
$SourceFolder = GUICtrlCreateInput($sread, 344, 112, 121, 21)
$SourceLabel = GUICtrlCreateLabel("Source Folder", 172, 120, 83, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$DestinationFolder = GUICtrlCreateInput($dread, 344, 168, 121, 21)
$DestinationLabel = GUICtrlCreateLabel("Destination Folder", 172, 176, 107, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$BackupLabel = GUICtrlCreateLabel("Backup Time", 172, 232, 78, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$BackupupTime = GUICtrlCreateInput($timeread, 344, 224, 121, 21)
$LogEnable = GUICtrlCreateCheckbox("Enable Logs", 168, 344, 97, 17)
$Silent = GUICtrlCreateCheckbox("Silent Mode", 168, 368, 97, 17)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Cancel
			Exit


		Case $Save
			$sreading = GUICtrlRead($SourceFolder)
			$dreading = GUICtrlRead($DestinationFolder)
			$backupTimer = GUICtrlRead($BackupupTime)



			$swrite = IniWrite("Config.ini","Section",'SourceFolder',$sreading)
			$dwrite = IniWrite("Config.ini","Section",'DestinationFolder',$dreading)
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



		$createShortcut = FileCreateShortcut($installdir & "InitialSetup.exe",@DesktopDir & "\OutlookAuto.lnk","","","OutlookAuto",$installdir & "OutlookBackup.ico")
		FileChangeDir($spath)
		$Enginestart = Run('Outlook_auto.exe')


		Exit

	EndSwitch
WEnd







