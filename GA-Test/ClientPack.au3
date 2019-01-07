#RequireAdmin
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <ColorConstants.au3>
#include <Color.au3>
#include <MsgBoxConstants.au3>



Global $arch = @OSArch
Global $spath = @ScriptDir
DirCreate($spath & "\Clientpack\")
Global $installdir = $spath & "\Clientpack\"
#CS

FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\Client.7z", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\7za.exe", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\extractor.exe", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\side.bmp", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\top3.bmp", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\top5.bmp", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\finalpage.exe", $installdir)
FileInstall("C:\Users\USER\Desktop\Test\EN-GUI\TestSuite.exe", $installdir)
#CE
;------------------------------------------------------------------- Getting IP And MaC Address of the system----------------------------------------------------------------

$IP_Address1 = @IPAddress1
$IP_Address2 = @IPAddress2
$IP_Address3 = @IPAddress3
$IP_Address4 = @IPAddress4

Global $IP_Address[4]

$IP_Address[0] = $IP_Address1
$IP_Address[1] = $IP_Address2
$IP_Address[2] = $IP_Address3
$IP_Address[3] = $IP_Address4


For $i = 0 To 3
	;MsgBox(0,"",$IP_Address[$i])
	If Not StringRegExp($IP_Address[$i], '(^127\.)|(^10\.)|(^172\.1[6-9]\.)|(^172\.2[0-9]\.)|(^172\.3[0-1]\.)|(^192\.168\.)') Then
							ContinueLoop
							;MsgBox(0, "MAC", "Invalid Private IPaddress")
	Else
							;MsgBox(0,"",$IP_Address[$i])
							Global $IP_AddressF =  $IP_Address[$i]

	EndIf

Next

;$IP_Address = @IPAddress1

Global $MAC_Address = GET_MAC($IP_AddressF)

;MsgBox(0, "MAC Address:", $MAC_Address)

Func GET_MAC($_MACsIP)
    Local $_MAC,$_MACSize
    Local $_MACi,$_MACs,$_MACr,$_MACiIP
    $_MAC = DllStructCreate("byte[6]")
    $_MACSize = DllStructCreate("int")
    DllStructSetData($_MACSize,1,6)
    $_MACr = DllCall ("Ws2_32.dll", "int", "inet_addr", "str", $_MACsIP)
    $_MACiIP = $_MACr[0]
    $_MACr = DllCall ("iphlpapi.dll", "int", "SendARP", "int", $_MACiIP, "int", 0, "ptr", DllStructGetPtr($_MAC), "ptr", DllStructGetPtr($_MACSize))
    $_MACs  = ""
    For $_MACi = 0 To 5
    If $_MACi Then $_MACs = $_MACs & ":"
        $_MACs = $_MACs & Hex(DllStructGetData($_MAC,1,$_MACi+1),2)
    Next
    DllClose($_MAC)
    DllClose($_MACSize)
    Return $_MACs
EndFunc

;-----------------------------------------------------------------------------------------------------------------------------------------------


#Region ### START Koda GUI section ### Form=
$Form1_1 = GUICreate("Client Setup Installer", 552, 418, -1, -1)
$Nextbutton = GUICtrlCreateButton("Next", 408, 352, 75, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$CancelButton = GUICtrlCreateButton("Cancel", 320, 352, 75, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label3 = GUICtrlCreateLabel("1. Python 2.7.14", 128, 176, 98, 26)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("2. Pywinauto 0.5.4", 128, 216, 148, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label5 = GUICtrlCreateLabel("3. WinPcap 4.1.3", 128, 256, 102, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label6 = GUICtrlCreateLabel("Click On Next To Continue", 128, 360, 155, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Pic1 = GUICtrlCreatePic($installdir & "\side.bmp", 0, -8, 76, 444)
$Pic2 = GUICtrlCreatePic($installdir & "\top3.bmp", 80, 0, 473, 65)
$Label2 = GUICtrlCreateLabel("The following components will be installed", 104, 128, 365, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $CancelButton
			Exit

		Case $Nextbutton
			GUIDelete($Form1_1)

					$Form1_2 = GUICreate("Initial Configuration", 552, 418, -1, -1)
					GUISetFont(10, 400, 0, "Times New Roman")
					$Label2 = GUICtrlCreateLabel("System IP Address", 128, 141, 121, 19)
					$Input1 = GUICtrlCreateInput($IP_AddressF, 256, 136, 169, 23)
					$Label3 = GUICtrlCreateLabel("System MAC Address", 112, 221, 140, 19)
					$Input2 = GUICtrlCreateInput($MAC_Address, 256, 216, 169, 23)
					$InitialConfigSaveBtn = GUICtrlCreateButton("Continue", 424, 373, 75, 25)
					$Initial_Config_Cancel = GUICtrlCreateButton("Cancel", 336, 373, 75, 25)
					$Pic1 = GUICtrlCreatePic($installdir & "\side.bmp", 0, -3, 76, 444)
					$Pic2 = GUICtrlCreatePic($installdir & "\top5.bmp", 80, -3, 449, 73)
					GUISetState(@SW_SHOW)
					#EndRegion ### END Koda GUI section ###

					While 1
						$nMsg = GUIGetMsg()
						Switch $nMsg
							Case $GUI_EVENT_CLOSE
								Exit

							Case $Initial_Config_Cancel
								Exit

							Case $InitialConfigSaveBtn
								$IPformUser = GUICtrlRead($Input1)
								$MACfromUser = GUICtrlRead($Input2)

								If Not $IPformUser Then
							MsgBox("Alert", "", "Please Enter IP address")
							;$MACfromUser = GUICtrlRead($Input2)
						ElseIf Not $MACfromUser Then
							MsgBox("Alert", "", "Please Enter MAC address")

							;IP Validation
						ElseIf Not StringRegExp($IPformUser, '^(?:(?:2(?:[0-4]\d|5[0-5])|1?\d{1,2})\.){3}(?:(?:2(?:[0-4]\d|5[0-5])|1?\d{1,2}))$') Then
							MsgBox(0, "MAC", "Please Enter a Valid IP address")
							;MAC Validation
						ElseIf Not StringRegExp($MACfromUser, "(?i)^[0-9A-F]{2}:[0-9A-F]{2}:[0-9A-F]{2}:[0-9A-F]{2}:[0-9A-F]{2}:[0-9A-F]{2}$") Then
							MsgBox(0, "MAC", "Please Enter a Valid MAC address")

						Else
							#CS
							FileChangeDir($installdir)
							$savelocation = FileOpen("Client.ini", 2)
							FileWrite($savelocation, "[Server]" & @CRLF)
							FileWrite($savelocation, "ServerIP=172.17.10.143" & @CRLF)
							FileWrite($savelocation, "ServerPort=5555" & @CRLF)
							FileWrite($savelocation, "[Client]" & @CRLF)
							FileWrite($savelocation, "ClientIP=" & $IP_AddressF & @CRLF)
							FileWrite($savelocation, "MAC=" & $MACfromUser & @CRLF)
							FileWrite($savelocation, "[Test]" & @CRLF)
							FileWrite($savelocation, "hips=0" & @CRLF)
							FileClose($savelocation)
							MsgBox(0, "Info", "Setting Saved Sucessfully")
							GUIDelete($Form1_2)
							Run("extractor.exe")
							Exit
							#CE
						EndIf


						EndSwitch
					WEnd



	EndSwitch
WEnd
