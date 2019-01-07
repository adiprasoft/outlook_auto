
#include<Array.au3>

#CS
Func Def_Object()

   Global $objOutlook = ObjCreate("Outlook.Application")
   Global $objNamespace = $objOutlook.GetNamespace("MAPI")
   Global $strStoreName = "prash.roy@live.com"
;~    MsgBox(0,0,$strStoreName)
   Global $objStore = $objNamespace.Stores.Item($strStoreName)
   Global $objRoot = $objStore.GetRootFolder()
   Global $objInbox = $objRoot.folders("Inbox")

EndFunc


;MsgBox(0,0,$objOutlook)
Sleep(5000)
Def_Object()
Sleep(5000)
$objOutlook.quit
MsgBox(0,0,'Used Quit Method')
Sleep(5000)
MsgBox(0,0,'Using Null')
Global $objOutlook = Null
Global $objNamespace = Null
Global $strStoreName = Null
Global $objStore = Null
Global $objRoot = Null
Global $objInbox = Null
Sleep(5000)
MsgBox(0,0,'Used Null')


MsgBox(0,0,'Quit')


#CE

Global $Include_list[1][1] = [['Inbox']]
_ArrayDisplay($Include_list)
Func ArrayAdd($subfolder)
	;$Include_list[0][0] +=1
   _ArrayAdd($Include_list, $subfolder)

EndFunc

ArrayAdd('Sent Items')
_ArrayDisplay($Include_list)