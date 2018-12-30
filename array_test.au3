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

	for $subfolder in $objfolder.folders
		_ArrayAdd($IncludeList,$objfolder.Name & '\' & $subFolder.Name)
		$IncludeList[0] +=1
	Next
Next


_ArrayDisplay($IncludeList)

