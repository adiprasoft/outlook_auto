#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <Date.au3>
#include <File.au3>

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
   Sleep(5000)
   $objOutlook.Quit()
   MsgBox(0,0,"Quit")
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

Dim $objSubfolders
For $objSubfolders in $objfolder.folders
	MsgBox(0,0,$objSubfolders.Name)
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
	    $fullpath = $DestFolder & "\"; & $objWorkingFolder & "\"
	    $DestFolder = StringReplace($fullpath, "\\", "\")

	    ;If Not FileExists($DestFolder & "\" & $objWorkingFolder) Then DirCreate($DestFolder & "\" & $objWorkingFolder)
		If Not FileExists($DestFolder) Then DirCreate($DestFolder)
;~ 	   MsgBox(0,'Final Folder', $final_folder)
	    SendAndLog("[Backup_mails]: Mail will be saved as " & $DestFolder & $filename & ".msg")
	    $mail.SaveAs($DestFolder & $filename & ".msg", 3)


	    If $delete_mails Then
		    MsgBox(0,'Delete Mail Value',$delete_mails)
		    $mail.Delete
	    EndIf

    Next

EndFunc



;_ArrayDisplay($IncludeList)



for $i = 1 to $IncludeList[0]
	Global $objfolder = $objRoot.folders($IncludeList[$i])

	;MsgBox(0,0,$objfolder.Name)
	Global $delete_mails = 0
	$x = _NowDate()
    SendAndLog("[Main]: Now Date" & _NowDate)
    $y = StringReplace($x, '/', '-')
    SendAndLog("[Main]: Now Date rectified" & $y)
    Global $DestFolder = $dst & "\" & $y & "\" & $objfolder.Name
	Backup_mails($DestFolder, $objfolder, $delete_mails)


Next


