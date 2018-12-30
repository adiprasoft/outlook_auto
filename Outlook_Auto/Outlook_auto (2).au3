#RequireAdmin
#include <Date.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include<Array.au3>
#include <Date.au3>

Global $DOS, $Message, $src, $dst, $bkpTime, $enablebkp, $logEnable = ""
$sKeyPath = "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook"
$cfgFile = @ScriptDir & "\Config.ini"
$log = False

Func SendAndLog($Data, $FileName = -1, $TimeStamp = True, $Log = $logEnable)
   If $Log Then
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

Func get_outlook_pst()
   ; Get Path of pst files from registry for all outlook accounts
   Dim $pst_paths[1] = [0]
   For $i= 1 to 50
	   $sKey = RegEnumKey($sKeyPath, $i)
		   If @error <> 0 then ExitLoop
	   For $ii = 1 to 100
		   $sEnumPath = $sKeyPath & "\" & $sKey
		   ; ConsoleWrite($sEnumPath & @LF)
			   $sKeyVal = RegEnumVal($sEnumPath, $ii)

		   If @error <> 0 Then ExitLoop
		   ; ConsoleWrite("Value Name  #" & $ii & " under in  key " & $sKeyVal & @LF)
		   $sRegVal = RegRead ($sEnumPath, $sKeyVal)
			   If Bitor(StringRight($sRegVal,4) = ".PST", StringRight($sRegVal,4) = ".OST") then
				   ;ConsoleWrite("SubKey #" & $i & " under HKCU: " & $sKey & @LF)
				   ;ConsoleWrite("PST PATH : " & $sRegVal & @LF)
				   _ArrayAdd($pst_paths, $sRegVal)
				   SendAndLog("[get_outlook_pst] Got pst path as : " & $sRegVal )
				   $pst_paths[0] += 1
			   EndIf
	   next
	Next
	Return $pst_paths
 EndFunc

Func folder_copy($src, $dst)
   ; Folder copy and create destination path if not present
   FileCopy($src & "\*", $dst & "\*", 8)
  ; SendAndLog("[folder_copy] FileCopy from : " & $src & "\*" & " to " $dst & "\*" & "And Error is " & @error)
   Return @error
EndFunc

Func ConfigCreate($cfgFile)
   ;Create config file for settings

   If FileExists($cfgFile) Then
	  SendAndLog("[ConfigCreate] Config file exists.")
	  return True
   Else
	  SendAndLog("[ConfigCreate] Config file does not exists. Creating a new config file")
	  Dim $data[7][2] = [[1, ""], ["SourceFolder", "Auto"], ["DestinationFolder", "D:\"], ["BackupTime", "19:00:00"], ["Reschedule","False"], ["EnableBackup", "True"], ["LogEnable", "False"]]
	  IniWriteSection($cfgFile, "Section", $data)
   EndIf
EndFunc

Func ConfigRead()
   ; Get settings from Config.ini file
   If Not FileExists($cfgFile) Then
	  SendAndLog("[ConfigRead] Config file does not exists.")
	  ConfigCreate($cfgFile)
   Else
	  Global $src = IniRead($cfgFile, "Section", "SourceFolder", "Auto")
	  Global $dst = IniRead($cfgFile, "Section", "DestinationFolder", "D:\")
	  Global $bkpTime = IniRead($cfgFile, "Section", "BackupTime", "19:00:00")
	  Global $reschedule = IniRead($cfgFile, "Section", "Reschedule", "False")
	  Global $enablebkp = IniRead($cfgFile, "Section", "EnableBackup", "True")
	  Global $logEnable = IniRead($cfgFile, "Section", "LogEnable", "False")
   EndIf
   Return True
EndFunc

Func Task_schedule_check()
   ; Check task in task schedular with name OutlookAuto
   If $reschedule Then
	  SendAndLog("[Task_schedule_check] Reschedule check enabled. Rescheduling task.")
	  task_schedule_delete()
   EndIf
   $DOS = Run(@ComSpec & " /c schtasks /Query /TN OutlookAuto", "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
   ProcessWaitClose($DOS)
   $Message = StdoutRead($DOS)
   SendAndLog("[Task_schedule_check] task query $Message = " & $Message)
   If Not StringInStr($Message, "OutlookAuto") Then task_schedule_create()
EndFunc

Func task_schedule_create()
   ;Create OutlookAuto task in task schedular
   SendAndLog("[task_schedule_create]: Creating task OutlookAuto")
   $DOS = Run(@ComSpec & ' /c SchTasks /Create /SC DAILY /TN “OutlookAuto” /TR “' & @ScriptDir & "\Outlook_auto.exe" & '” /ST '& $bkpTime & '"', "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
   ProcessWaitClose($DOS)
EndFunc

Func task_schedule_delete()
   SendAndLog("[task_schedule_delete]: Deleting task OutlookAuto")
   $DOS = Run(@ComSpec & ' /c schtasks /delete /tn OutlookAuto /f', "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
   ProcessWaitClose($DOS)
   SendAndLog("[task_schedule_delete]: reseting config")
   IniWrite($cfgFile, "Section", "Reschedule", "False")
EndFunc

Func Outlook_Launch_exit()
   SendAndLog("[Outlook_Launch_exit]: running outlook")
   $outlook = Run("outlook.exe")
   SendAndLog("[Outlook_Launch_exit]: outlook pid " & $outlook)
   Sleep(1000*60)
   ProcessClose($outlook)
EndFunc

Func autologon()
   If RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon") == 0 Then
	  SendAndLog("[autologon]: Winlogon 32 bit is 0")
	  RunWait(@ScriptDir & "\Autologon.exe")
   EndIf
   If RegRead("HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon") == 0 Then
	  SendAndLog("[autologon]: Winlogon 64Bit is 0")
	  RunWait(@ScriptDir & "\Autologon.exe")
   EndIf
   ;MsgBox(0,"",@ScriptDir & "\Autologon.exe")
EndFunc

Func Main()
   SendAndLog("[Main]: Start")
   SendAndLog("[Main]: Sleeping for some time")
   Sleep(1000*60)
   SendAndLog("[Main]: Resuming")
   ConfigRead() ; Read Config File
   autologon()  ; Check autologon

   If $enablebkp Then
	  Task_schedule_check()
   Else
	  task_schedule_delete()
   EndIf

   Outlook_Launch_exit()
   If StringCompare($src , "Auto") == 0 Then Global $src = get_outlook_pst()
   If Not IsArray($src) Then
	  $DestFolder = $dst & "\" & _NowDate() & "\"
	  SendAndLog("[Main]: Copying Folder")
	  folder_copy($src, $DestFolder)
   Else
	  for $i =1 to $src[0]
		 SendAndLog("[Main]: Copying Folders, ", $src[$i])
		 folder_copy($src[$i], $DestFolder)
	  Next
   EndIf
EndFunc

;Func Test()
;   autologon()
;EndFunc

FileInstall("D:\Scripts\Outlook_Auto\Autologon.exe", @scriptdir & "\Autologon.exe")
Main()
;Test()
