

#include <Array.au3>
#include <MsgBoxConstants.au3>

Local $aArray[20]
For $i = 0 To 19
    $aArray[$i] = $i
Next
_ArrayDisplay($aArray, "1D Array")
MsgBox($MB_SYSTEMMODAL, "Items 1-7", _ArrayToString($aArray, ','))

#CS
Local $aArray[10][10]
For $i = 0 To 9
    For $j = 0 To 9
        $aArray[$i][$j] = $i & "-" & $j
    Next
Next
_ArrayDisplay($aArray, "2D Array")
MsgBox($MB_SYSTEMMODAL, "Rows 4-7,  cols 2-5", _ArrayToString($aArray, " :: ", 4, 7, @CRLF, 2, 5))
#CE


Dim $Include_list[1]
$Include_list[0] = "Sent items"
_ArrayAdd($Include_list, "Inbox")

MsgBox(0,0,_ArrayToString($Include_list,','))

_ArrayDisplay($Include_list)