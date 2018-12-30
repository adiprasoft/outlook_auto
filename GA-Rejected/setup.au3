#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=img\installer.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>

Global $arch = @OSArch
Global $spath = @ScriptDir

;MsgBox(0,0,$check)

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Installer", 630, 456, -1, -1)
$Company = GUICtrlCreatePic("C:\FreeLance\img\soft.bmp", -8, 56, 76, 398)
$Outlook = GUICtrlCreatePic("C:\FreeLance\img\try.jpg", 0, 0, 628, 60)
$Blank = GUICtrlCreatePic("C:\FreeLance\img\black.bmp", -8, 56, 636, 4)
$Label1 = GUICtrlCreateLabel("Install Outlook Backup Creator", 112, 128, 437, 40)
GUICtrlSetFont(-1, 23, 800, 0, "MS Sans Serif")
$Install = GUICtrlCreateButton("Install", 472, 352, 75, 25)
$Cancel = GUICtrlCreateButton("Cancel", 384, 352, 75, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Cancel
			Exit

		Case $Install
			DirCreate("C:\OutlookAuto")
			Global $installdir = "C:\OutlookAuto\"
			;MsgBox(0,0,$installdir)
			$outlook = FileInstall("C:\FreeLance\Outlook_auto.exe",$installdir,1)
			$config = FileInstall("C:\FreeLance\Config.ini",$installdir,1)
			$autologon = FileInstall("C:\FreeLance\Autologon.exe",$installdir,1)
			$InitialSet = FileInstall("C:\FreeLance\InitialSetup.exe",$installdir,1)
			$shortcut = FileInstall("C:\FreeLance\img\OutlookBackup.ico",$installdir,1)
			GUIDelete($Form1)



			#Region ### START Koda GUI section ### Form=
			$Installing = GUICreate("Installing Setup", 630, 456, -1, -1)
			$Company = GUICtrlCreatePic("C:\FreeLance\img\soft.bmp", -8, 56, 76, 398)
			$Outlook = GUICtrlCreatePic("C:\FreeLance\img\try.jpg", 0, 0, 628, 60)
			$Blank = GUICtrlCreatePic("C:\FreeLance\img\black.bmp", -8, 56, 636, 4)
			$ProgressBar = GUICtrlCreateProgress(112, 168, 478, 17)
			$Installing = GUICtrlCreateLabel("Installing...", 112, 112, 130, 36)
			;ProgressOn
			GUICtrlSetFont(-1, 20, 800, 0, "MS Sans Serif")
			GUISetState(@SW_SHOW)

			$timer = TimerInit()
			#EndRegion ### END Koda GUI section ###
			$time = 1000

			While 1
				if GUIGetMsg() = $GUI_EVENT_CLOSE then Exit
				$progress_data = GUICtrlRead($progressbar)
				if $progress_data = 100 Then
					Sleep(600)
					;MsgBox(0,"Status","Installation Complete", 10)
					Sleep(1000)
					$InitialSetup = Run($installdir & 'InitialSetup.exe')
					Exit

				EndIf
				$Percentage_data = TimerDiff($timer)*100/$time
				GUICtrlSetData($progressbar,$Percentage_data)
			WEnd

	EndSwitch

WEnd
