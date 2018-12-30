#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
Global $Form1 = GUICreate("Form1", 616, 439, 192, 124)
GUICtrlCreateGroup("Group1", 246, 91, 254, 186)
Global $DummyStart = GUICtrlCreateDummy() ; get start of control creation control id
GUICtrlCreateRadio("Radio1", 312, 108, 113, 17)
GUICtrlCreateRadio("Radio2", 312, 231, 113, 23)
GUICtrlCreateRadio("Radio3", 312, 211, 113, 17)
GUICtrlCreateRadio("Radio4", 312, 149, 113, 17)
GUICtrlCreateRadio("Radio5", 312, 190, 113, 17)
GUICtrlCreateRadio("Radio6", 312, 129, 113, 17)
GUICtrlCreateRadio("Radio7", 312, 169, 113, 17)
Global $DummyEnd = GUICtrlCreateDummy() ; get end of control creation id
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $Button1 = GUICtrlCreateButton(" Hide ", 273, 372, 75, 25)
GUISetState(@SW_SHOW)

While 1
    Sleep(100)
    $msg = GUIGetMsg()
    Switch $msg
        Case $Button1
            Button1Click()
        Case $GUI_Event_Close
            Exit
    EndSwitch
WEnd

Func Button1Click()
    Local Static $toggle = True
    $toggle = Not $toggle
    For $Loop = $DummyStart + 1 To $DummyEnd - 1
        If $toggle Then
            GUICtrlSetState($Loop, $GUI_SHOW)
            GUICtrlSetData($Button1, " Hide ")
        Else
            GUICtrlSetState($Loop, $GUI_HIDE)
            GUICtrlSetData($Button1, " Show ")
        EndIf
    Next
EndFunc   ;==>Button1Click