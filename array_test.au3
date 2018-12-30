#include<array.au3>

$strStoreName='prash.roy@live.com'

Global $objOutlook = ObjCreate("Outlook.Application")
Global $objNamespace = $objOutlook.GetNamespace("MAPI")
;Global $strStoreName = IniRead($cfgFile,"Section","MailID","")

Global $objStore = $objNamespace.Stores.Item($strStoreName)
Global $objRoot = $objStore.GetRootFolder()


Global $IncludeList = StringSplit(IniRead('Config.ini', "Section", "Outlook_Folder_Include_List", ""), ",")
;_ArrayDisplay($IncludeList,'Include List')

for $i = 1 to $IncludeList[0]
	Global $objfolder = $objRoot.folders($IncludeList[$i])
	MsgBox(0,0,$objfolder.Name)
	Global $subFolderArray[1] = [0]
	for $subfolder in $objfolder.folders
		;MsgBox(0,'Subfolder',$subfolder.Name)
		_ArrayAdd($subFolderArray,$objfolder.Name & '\' & $subFolder.Name)
		$subFolderArray[0] += 1
		_ArrayDisplay($subFolderArray)
	Next
Next

_ArrayAdd($IncludeList,$subFolderArray)
_ArrayDisplay($IncludeList)

