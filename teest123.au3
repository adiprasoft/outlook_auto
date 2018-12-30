#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include<Array.au3>
#include <Date.au3>

Global $cfgFile = @ScriptDir & "\Config.ini"
Global $DestFolder = "D:\Outlook_Backup"
$dst = $DestFolder


Global $comError = ObjEvent("Autoit.Error","MyErrFunc")	;Initialize a COM error Handler
Global $comErrorCount = 0


Func MyErrFunc($comError)
	$comErrorCount += 1
	SendAndLog("[COM_ERROR]: COM ERROR INTERCEPTED" & " " & $comErrorCount)
EndFunc


Func Def_Object()
   Global $objOutlook = ObjCreate("Outlook.Application")
   Global $objNamespace = $objOutlook.GetNamespace("MAPI")
   Global $strStoreName = IniRead($cfgFile,"Section","MailID","")
;~    MsgBox(0,0,$strStoreName)
   Global $objStore = $objNamespace.Stores.Item($strStoreName)
   Global $objRoot = $objStore.GetRootFolder()
   ;Global $objInbox = $objRoot.folders("Inbox")
EndFunc

Def_Object()

Func SendAndLog($Data, $FileName = -1, $TimeStamp = True, $Log = True)
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



Global $IncludeList = StringSplit(IniRead('Config.ini', "Section", "Outlook_Folder_Include_List", ""), ",")
;_ArrayDisplay($IncludeList,'Include List')

;================================================
;Function to add subfolders
;================================================

for $i = 1 to $IncludeList[0]
	Global $objfolder = $objRoot.folders($IncludeList[$i])

	for $subfolder in $objfolder.folders
		_ArrayAdd($IncludeList,$objfolder.Name & '\' & $subFolder.Name)
		$IncludeList[0] +=1
	Next
Next


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
	$filename = StringReplace($filename, "´",   "'")
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


Func Backup_mails($DestFolder, $objfolder, $delete_mails = 0)

	;MsgBox(0,'Delete Mail Value',$delete_mails)
    $counter = $objfolder.items.count
    $colitems = $objfolder.items
    $objWorkingFolder = $objfolder.Name
    for $mail in $colitems

	    $filename = $mail.subject & "-" & $mail.ReceivedTime
	    SendAndLog("[Backup_mails]: File Name 1 - " & $filename)

		$filename = CleanString($filename)

	    SendAndLog("[Backup_mails]: File Name 2 - " & $filename)
	    $final_folder = $DestFolder & "\" & $objWorkingFolder & "\"
	    $final_folder = StringReplace($final_folder, "\\", "\")

	    If Not FileExists($DestFolder & "\" & $objWorkingFolder) Then DirCreate($DestFolder & "\" & $objWorkingFolder)
;~ 	   MsgBox(0,'Final Folder', $final_folder)
	    SendAndLog("[Backup_mails]: Mail will be saved as " & $final_folder & $filename & ".msg")
	    $mail.SaveAs($final_folder & $filename & ".msg", 3)


	    If $delete_mails Then
		    MsgBox(0,'Delete Mail Value',$delete_mails)
		    $mail.Delete
	    EndIf

    Next

EndFunc



_ArrayDisplay($IncludeList)

for $i = 1 to $IncludeList[0]
	Global $objfolder = $objRoot.folders($IncludeList[$i])
	;MsgBox(0,0,$objfolder.Name)
	Global $delete_mails = 0
	$x = _NowDate()
    SendAndLog("[Main]: Now Date" & _NowDate)
    $y = StringReplace($x, '/', '-')
    SendAndLog("[Main]: Now Date rectified" & $y)
    $DestFolder = $dst & "\" & $y & "\"
	Backup_mails($DestFolder, $objfolder, $delete_mails)
Next

$objOutlook.Destroy
