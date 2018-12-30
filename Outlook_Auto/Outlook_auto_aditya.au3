#RequireAdmin
#include <Date.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include<Array.au3>

Global $DOS, $Message, $src, $dst, $bkpTime, $enablebkp, $logEnable = ""
$sKeyPath = "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook"
$cfgFile = @ScriptDir & "\Config.ini"
$log = False
Global $logEnable = "False"
If FileExists(@scriptdir & "\Config.ini") Then
	  Global $logEnable =  IniRead($cfgFile, "Section", "LogEnable", "False")
EndIf

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

Func Sysinfo()
   SendAndLog("[Sysinfo]: Start")
   SendAndLog("[Sysinfo]: Operating System   : " & @OSBuild)
   SendAndLog("[Sysinfo]: Service Pack       : "  & @OSServicePack)
   SendAndLog("[Sysinfo]: System Architecture: " &  @CPUArch)
   SendAndLog("[Sysinfo]: Host               : "  & @ComputerName)
EndFunc

Func get_outlook_pst()
   SendAndLog("[get_outlook_pst]" )
   ; Get Path of pst files from registry for all outlook accounts
  	Dim $pst_paths[1] = [0]
	$vApp = ObjCreate("Outlook.Application")
	SendAndLog("[get_outlook_pst] : " & @error )
	For $vStore in $vApp.Session.Stores
	   SendAndLog("[get_outlook_pst]:" & $vStore.FilePath)
	   If StringInStr($vStore.FilePath, ".ost") Then
		  _ArrayAdd($pst_paths, $vStore.FilePath)
		  $pst_paths[0] += 1
		 SendAndLog("[get_outlook_pst]:  $pst_paths[0] = " &  $pst_paths[0])
	  EndIf
   Next
   SendAndLog("[get_outlook_pst]: Returned")
   Return $pst_paths
 EndFunc

Func folder_copy($src, $dst)
   ; Folder copy and create destination path if not present
   FileMove($src & "\*", $dst & "\*", 8)
   SendAndLog("[folder_copy] FileCopy from : " & $src & "\*" & " to " & $dst & "\*" & "And Error is " & @error)
   Return @error
EndFunc

Func ConfigCreate($cfgFile)
   ;Create config file for settings
   If FileExists($cfgFile) Then
	  SendAndLog("[ConfigCreate] Config file exists.")
	  return True
   Else
	  SendAndLog("[ConfigCreate] Config file does not exists. Creating a new config file")
	  Dim $data[7][2] = [[1, ""], ["SourceFolder", "C:\Data"], ["DestinationFolder", "D:\Outlook_backup"], ["BackupTime", "19:00:00"], ["Reschedule","False"], ["EnableBackup", "True"], ["LogEnable", "False"]]
	  IniWriteSection($cfgFile, "Section", $data)
   EndIf
EndFunc

Func ConfigRead()
   ; Get settings from Config.ini file
   If Not FileExists($cfgFile) Then
	  SendAndLog("[ConfigRead] Config file does not exists.")
	  ConfigCreate($cfgFile)
   Else
	  Global $src = IniRead($cfgFile, "Section", "SourceFolder", "C:\Data")
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
   SendAndLog("[Task_schedule_check]: Reschedule check, " & $reschedule)
   If StringCompare($reschedule, 'True')==0 Then
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
   If StringInStr(@ScriptDir & '\Outlook_auto.exe', ' ') <> 0 Then
	  $cmd = ' /c SchTasks /Create /SC DAILY /TN "OutlookAuto /TR" & @ScriptDir & "\Outlook_auto.exe" & " /ST "& $bkpTime '
   Else
	  $cmd = ' /c SchTasks /Create /SC DAILY /TN OutlookAuto /TR  & @ScriptDir & "\Outlook_auto.exe" & " /ST "& $bkpTime '
   EndIf
   $DOS = Run(@ComSpec & $cmd, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
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
   $outlook_path = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe", "Path")
   SendAndLog("[Outlook_Launch_exit]: Outlook 32 " & $outlook_path)
   If Not $outlook_path Then
	  $outlook_path = RegRead("HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe", "Path")
	  SendAndLog("[Outlook_Launch_exit]: Outlook 64 " & $outlook_path)
	  If Not $outlook_path Then SendAndLog("[Outlook_Launch_exit]: Outlook Not found " & $outlook_path)
   EndIf
   SendAndLog("[Outlook_Launch_exit]: running outlook")
   $outlook = Run($outlook_path & "\outlook.exe")
   SendAndLog("[Outlook_Launch_exit]: outlook pid " & $outlook)
   For $i=30 to 0 Step -1
	  TrayTip("OutlookAuto", "Waiting " & $i &" mins for mail to dowload", 5)
	  Sleep(1000*60)
   Next
   If $outlook <> 0 Then ProcessClose($outlook)
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
   Sysinfo()
   SendAndLog("[Main]: Start")
   SendAndLog("[Main]: Sleeping for some time")
   TrayTip("OutlookAuto", "Running outlook in 60 sec", 5)
   Sleep(1000*60)
   SendAndLog("[Main]: Resuming")
   ConfigRead() ; Read Config File
   autologon()  ; Check autologon

   If StringCompare($enablebkp, 'True')==0  Then
	  Task_schedule_check()
   Else
	  task_schedule_delete()
   EndIf

   Outlook_Launch_exit()

   If StringCompare($src , "Auto") == 0 Then Global $src = get_outlook_pst()
   SendAndLog("[Main]: Source isArray" & IsArray($src))
   If Not IsArray($src) Then
	  SendAndLog("[Main]: Not an array")
	  $x = _NowDate()
	  SendAndLog("[Main]: Now Date" & _NowDate)
	  $y = StringReplace($x, '/', '-')
	  SendAndLog("[Main]: Now Date rectified" & $y)
	  $DestFolder = $dst & "\" & $y & "\"
	  SendAndLog("[Main]: Copying Folder" & $src & $DestFolder)
	  folder_copy($src, $DestFolder)
   Else
	  SendAndLog("[Main]: Source is an array")
	  for $i =1 to $src[0]
		 SendAndLog("[Main]: Copying Folders, ", $src[$i])
		 folder_copy($src[$i], $DestFolder)
	  Next
   EndIf
EndFunc

FileInstall("C:\FreeLance\Autologon.exe", @scriptdir & "\Autologon.exe")
Main()