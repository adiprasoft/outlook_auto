#RequireAdmin
#include <Date.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include<Array.au3>


Global $DOS, $Message, $dst, $bkpTime,  $logEnable = ""
$sKeyPath = "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook"
$cfgFile = @ScriptDir & "\Config.ini"
$log = False
Global $logEnable = "False"
Global $objArray[1][2] = [[0,0]]

If FileExists(@scriptdir & "\Config.ini") Then
	  Global $logEnable =  IniRead($cfgFile, "Section", "LogEnable", "False")
EndIf

Global $comError = ObjEvent("Autoit.Error","MyErrFunc")	;Initialize a COM error Handler
Global $comErrorCount = 0


Func MyErrFunc($comError)
	$comErrorCount += 1
	SendAndLog("[COM_ERROR]: COM ERROR INTERCEPTED" & " " & $comErrorCount)
EndFunc


Func Set_startup_run()
   SendAndLog("[Set_startup_run]: Start")
   $value = RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "OutlookAuto")
   SendAndLog("[Set_startup_run]: Registry values is - " & $value)
   If $value = "" Then
	  SendAndLog("[Set_startup_run]: Creating registry entry for startup - " & @ScriptDir & "\Outlook_auto.exe")
	  RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "OutlookAuto", "REG_SZ", @ScriptDir & "\Outlook_auto.exe")
   EndIf
EndFunc


Func Sysinfo()
   SendAndLog("[Sysinfo]: Start")
   SendAndLog("[Sysinfo]: Operating System   : " & @OSBuild)
   SendAndLog("[Sysinfo]: Service Pack       : "  & @OSServicePack)
   SendAndLog("[Sysinfo]: System Architecture: " &  @CPUArch)
   SendAndLog("[Sysinfo]: Host               : "  & @ComputerName)
EndFunc



;===========================================================
;Create Outlook Object
;===========================================================
Func Def_Object()
   SendAndLog("[Def_Object]: Start")
   Global $objOutlook = ObjCreate("Outlook.Application")
   SendAndLog("[Def_Object]: Outlook.Application Created")
   Global $objNamespace = $objOutlook.GetNamespace("MAPI")
   SendAndLog("[Def_Object]: NameSpace Created")
   Global $strStoreName = IniRead($cfgFile,"Section","MailID","")
;~    MsgBox(0,0,$strStoreName)
   Global $objStore = $objNamespace.Stores.Item($strStoreName)
   SendAndLog("[Def_Object]: StoreName Created")
   Global $objRoot = $objStore.GetRootFolder()
   SendAndLog("[Def_Object]: $objRoot = " & $objRoot )
  Global $objInbox = $objRoot.folders("Inbox")
  SendAndLog("[Def_Object]: Obj Inbox Created")
EndFunc


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


Func ConfigCreate($cfgFile)
   ;Create config file for settings
   If FileExists($cfgFile) Then
	  SendAndLog("[ConfigCreate] Config file exists.")
	  return True
   Else
	  SendAndLog("[ConfigCreate] Config file does not exists. Creating a new config file")
	  ;Dim $data[7][2] = [[1, ""], ["DeleteMails", 0], ["DestinationFolder", "D:\Outlook_backup"], ["BackupTime", "19:00:00"], ["LogEnable", "False"],["Outlook_Folder_Include_List", ""],["MailID",""]]
	  Dim $data[6][2] = [["DeleteMails", 0], ["DestinationFolder", "D:\Outlook_backup"], ["BackupTime", "19:00:00"], ["LogEnable", "False"],["Outlook_Folder_Include_List", ""],["MailID",""]]
	  IniWriteSection($cfgFile, "Section", $data)
   EndIf
EndFunc



;===========================================================
;Read or Create Config file
;===========================================================

Func ConfigRead()
   ; Get settings from Config.ini file
   If Not FileExists($cfgFile) Then
	  SendAndLog("[ConfigRead] Config file does not exists.")
	  ConfigCreate($cfgFile)
   Else
	  Global $dst = IniRead($cfgFile, "Section", "DestinationFolder", "D:\Outlook_Backup") & "\"
	  Global $bkpTime = StringSplit(IniRead($cfgFile, "Section", "BackupTime", "19:00:00"), ":")
	  Global $mail_id = IniRead($cfgFile, "Section", "MailID", "abc.xyz.com")
	  Global $delete_mails = IniRead($cfgFile, "Section", "DeleteMails", 0)
	  Global $logEnable = IniRead($cfgFile, "Section", "LogEnable", "False")
	  Global $IncludeList = StringSplit(IniRead($cfgFile, "Section", "Outlook_Folder_Include_List", ""), ",")
	EndIf
   Return True
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


Func Download_Mails()
   SendAndLog("[Download_Mails]: Start")
   $count = 0
   Do
	   $org_cnt = $objInbox.items.count
		SendAndLog("[Download_Mails]: Mail count 1 - " & $org_cnt)
		If Not $count = $org_cnt Then
		   sleep(1000*30)
		EndIf
		$count = $objInbox.items.count
		SendAndLog("[Download_Mails]: Mail count 2 - " & $count)
		Until $count = $org_cnt
EndFunc


Func ArrayAdd($subArray)
   SendAndLog("[ArrayAdd]: Start")
   _ArrayAdd($objArray, $subArray)
   SendAndLog("[ArrayAdd]: Added " & $objArray & " and " & $subArray)
   $objArray[0][0] +=1
   SendAndLog("[ArrayAdd]: Array count - " & $objArray[0][0])
EndFunc


Func Get_Mail_folders_count()
   SendAndLog("[Get_Mail_folders_count]: Start")
   For $objSubFolder in $objRoot.folders
	   Dim $y[1][2] = [[$objSubFolder.Name, $objRoot.folders($objSubFolder.Name).items.count]]
	   _ArrayAdd($objArray, $y)
	   $objArray[0][0] += 1
	   SendAndLog("[Get_Mail_folders_count]: Folder Name - " & $objSubFolder.Name & "Mail Count - " & $objRoot.folders($objSubFolder.Name).items.count)
	Next
	SendAndLog("[Get_Mail_folders_count]: Total Folder Count - " & $objArray[0][0])

   Dim $objSubFolder
   For $objSubFolder in $objInbox.folders
	  SendAndLog("[Get_Mail_folders_count]: for $objSubFolder - " & $objSubFolder)
	   $cnt=$objSubFolder.items.count
	   SendAndLog("[Get_Mail_folders_count]: Count is - "& $cnt)
	   Global $subArray[1][2] = [[$objInbox.Name & '\' & $objSubFolder.Name, $cnt]]
	   ArrayAdd($subArray)
	Next
	SendAndLog("[Get_Mail_folders_count]: End")
	Return $objArray
EndFunc


Func Service_Wait()
   SendAndLog("[Service_Wait]: Start")
      If IsArray($bkpTime) Then
	  If $bkpTime[0] = 3 Then
		 While Not (@HOUR = $bkpTime[1] And $bkpTime[2] = @MIN)
		 Sleep (1000*5)
		 WEnd
	  EndIf
   EndIf
   SendAndLog("[Service_Wait]: End")
EndFunc

;===========================================================
;Clenaing Mail Subject Line
;===========================================================


Func CleanString($filename)
	$filename = StringReplace($filename, "/" , "")
	$filename =	StringReplace($filename, "\" , "")
	$filename =	StringReplace($filename, ":" , "")
	$filename =	StringReplace($filename, "*" , "")
	$filename =	StringReplace($filename, "?" , "")
	$filename =	StringReplace($filename, '"' , "")
	$filename =	StringReplace($filename, "|" , "")
	$filename =	StringReplace($filename, "^" , "")
	$filename =	StringReplace($filename, "." , "")
	$filename =	StringReplace($filename, "	" , " ") ; Added More
	$filename = StringReplace($filename, "   ", " ")
	$filename = StringReplace($filename, "�",   "'")
    $filename = StringReplace($filename, "`",   "'")
    $filename = StringReplace($filename, "{",   "(")
    $filename = StringReplace($filename, "[",   "(")
    $filename = StringReplace($filename, "]",   ")")
    $filename = StringReplace($filename, "}",   ")")
    $filename = StringReplace($filename, """",  "'")
    $filename = StringReplace($filename, "<",   "_")
    $filename = StringReplace($filename, ">",   "_")
	Return($filename)
EndFunc

;================================================
;Function to add subfolders
;================================================
Func Add_Subfolders($DestFolder, $objfolder)
	Dim $objSubfolders
	For $objSubfolders in $objfolder.folders
		;MsgBox(0,0,$objSubfolders.Name)
		$SubDestFolder = $DestFolder & '\' & $objSubfolders.Name
		Backup_mails($SubDestFolder, $objSubfolders)

		If $objSubfolders.folders.count > 1 Then
			Dim $fold
			For $fold in $objSubfolders.folders
				$DestFolder = $DestFolder & '\' & $objSubfolders & '\' & $fold.Name
				Backup_mails($DestFolder, $fold)
			Next
		EndIf
	Next
EndFunc



;===========================================================
; Copy and Delete mail Function
;===========================================================

Func Backup_mails($DestFolder, $objfolder, $delete_mails = 0)

    $counter = $objfolder.items.count
    $colitems = $objfolder.items
    $objWorkingFolder = $objfolder.Name
    for $mail in $colitems

	    $filename = $mail.subject & "-" & $mail.ReceivedTime
	    SendAndLog("[Backup_mails]: File Name 1 - " & $filename)

		$filename = CleanString($filename)

	    SendAndLog("[Backup_mails]: File Name 2 - " & $filename)
	    $fullpath = $DestFolder & "\"
	    $DestFolder = StringReplace($fullpath, "\\", "\")


		If Not FileExists($DestFolder) Then DirCreate($DestFolder)
	    SendAndLog("[Backup_mails]: Mail will be saved as " & $DestFolder & $filename & ".msg")

	    $mail.SaveAs($DestFolder & $filename & ".msg", 3)

	    If $delete_mails Then
			SendAndLog("[Deleting_Mails]: Mail will be deleted " & $DestFolder & $filename & ".msg")
			$mail.Delete
	    EndIf

    Next

EndFunc



;================================================================
; Main Function
;================================================================


Func Main()
	Dim $Mail_count
    TrayTip("Outlook Auto", "Starting in 30 sec", 5)
    Sleep(1000*25)
    SendAndLog("[Main]:")
    Sysinfo()
    ConfigRead()
    Set_startup_run()
    autologon()

   ;-----------------
    Def_Object()
    Download_Mails()
    $Mail_count = Get_Mail_folders_count()

	;===========================================================================
	; Creating $obj for Inbox,Del,sent etc
	;===========================================================================
	for $i = 1 to $IncludeList[0]
		Global $objfolder = $objRoot.folders($IncludeList[$i])
		$x = _NowDate()
		SendAndLog("[Main]: Now Date" & _NowDate)
		$y = StringReplace($x, '/', '-')
		SendAndLog("[Main]: Now Date rectified" & $y)
		Global $DestFolder = $dst & "\" & $y & "\" & $objfolder.Name
		$out = Backup_mails($DestFolder, $objfolder)
		SendAndLog("[Main]: Backup status" & $out)

		SendAndLog("[Nested_Folder]: Checking for Nested Folder Under " & $objfolder.Name)
		$Nested = Add_Subfolders($DestFolder, $objfolder)
		SendAndLog("[Nested_Folder]: Status for Nested Folder " & $Nested)
	Next
	TrayTip("Outlook Auto","Mail Backup Complete", 5)
EndFunc



While 1
	Main()
	Sleep(5000)
	MsgBox(0,"outlook pid",WinGetProcess("outlook"))
	ProcessClose(WinGetProcess("outlook"))
	Service_Wait()
WEnd