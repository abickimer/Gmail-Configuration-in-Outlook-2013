#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>

Opt("GUICoordMode", 2)
Opt('MustDeclareVars', 1)

AB1() ;Start Outlook and start manual config for new email account
AB2() ;Insert full name into new Account window
AB6() ;Insert email address
AB3() ;Insert incoming and outgoing Google mail servers (imap)
AB7() ;Insert email account user name
AB4() ;Insert email account password
AB9() ;Tab to password field
AB5() ;Receipt user password and input into outlook
AB11() ;insert port information

Func AB1()
Send("{LWIN}")
Send("cmd.exe")
Send("{ENTER}")
Sleep(1500)
Send("start outlook.exe")
Send("{ENTER}")
Sleep(10000)
Send("!f") ;alt + f
Send("{TAB 2}")
Send ("{ENTER}")
Send("m")
Send ("{ENTER}")
Send("p")
Send ("{ENTER}")
EndFunc

Func AB2()
    Local $Button, $Input, $msg, $iPID, $Form1
#Region ### START Koda GUI section ### Form=C:\Users\CyberPowerPC\Desktop\Outlook-scripts\koda\Form1.kxf
$Form1 = GUICreate("Form1", 615, 437, 192, 124)
$Input = GUICtrlCreateInput("Full Name", 56, 96, 505, 21)
$Button = GUICtrlCreateButton("Send to Outlook", -1, 136, 105, 65)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

    While 1
        $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                If ProcessExists($iPID) Then ProcessClose($iPID)
                ExitLoop
            Case $Button
                $iPID = WinActivate("Add Account")
                Sleep(250)
                Send(GUICtrlRead($Input))
				WinClose("[CLASS:AutoIt v3 GUI]")
        EndSwitch
    WEnd
 EndFunc

Func AB6()
   WinActivate("Add Account")
Send("{TAB}")
EndFunc

Func AB3()
    Local $Button, $Input, $msg, $iPID, $Form2
#Region ### START Koda GUI section ### Form=C:\Users\CyberPowerPC\Desktop\Outlook-scripts\koda\Form2.kxf
$Form2 = GUICreate("Form2", 745, 358, 302, 218)
$Button = GUICtrlCreateButton("Send to Outlook", 32, 168, 241, 65)
$Input = GUICtrlCreateInput("Email Address", -1, 64, 585, 21)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

    While 1
        $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                If ProcessExists($iPID) Then ProcessClose($iPID)
                ExitLoop
            Case $Button
                $iPID = WinActivate("Add Account")
                Sleep(250)
                Send(GUICtrlRead($Input))
				WinClose("[CLASS:AutoIt v3 GUI]")
        EndSwitch
    WEnd
 EndFunc

Func AB7()
   WinActivate("Add Account")
Send("{TAB}")
Send("i")
Send("{TAB}")
Send ("imap.googlemail.com"); input incoming mail server here
Send("{TAB}")
Send ("smtp.googlemail.com") ; input outgoing mail server here
Send("{TAB}")
EndFunc

Func AB4()
    Local $Button, $Input, $msg, $iPID, $Form3, $Label1
#Region ### START Koda GUI section ### Form=C:\Users\CyberPowerPC\Desktop\Outlook-scripts\koda\Form3.kxf
$Form3 = GUICreate("Form3", 747, 363, 302, 218)
$Input = GUICtrlCreateInput("UserName", 40, 40, 657, 21)
$Label1 = GUICtrlCreateLabel("(In most cases this is the same as your email address)", -1, 72, 254, 17)
$Button = GUICtrlCreateButton("Send to Outlook", -1, 64, 585, 49)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

    While 1
        $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                If ProcessExists($iPID) Then ProcessClose($iPID)
                ExitLoop
            Case $Button
                $iPID = WinActivate("Add Account")
                Sleep(250)
                Send(GUICtrlRead($Input))
				WinClose("[CLASS:AutoIt v3 GUI]")
        EndSwitch
    WEnd
 EndFunc

Func AB9 ()
   WinActivate("Add Account")
Send("{TAB}")
EndFunc

Func AB5()
    Local $Button, $Input, $msg, $iPID, $Form4
#Region ### START Koda GUI section ### Form=C:\Users\CyberPowerPC\Desktop\Outlook-scripts\koda\Form4.kxf
$Form4 = GUICreate("Form4", 750, 293, 302, 218)
$Input = GUICtrlCreateInput("Password", 32, 56, 673, 21)
$Button = GUICtrlCreateButton("Send to Outlook", -1, 104, 433, 65)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

    While 1
        $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                If ProcessExists($iPID) Then ProcessClose($iPID)
                ExitLoop
            Case $Button
                $iPID = WinActivate("Add Account")
                Sleep(250)
                Send(GUICtrlRead($Input))
				WinClose("[CLASS:AutoIt v3 GUI]")
        EndSwitch
	 WEnd
  EndFunc

Func AB11 ()
WinActivate("Add Account")
Send("{TAB}")
Send("m")
Send ("^{TAB}");cntrl + tab
Send ("{space}")
Send ("^{TAB}");cntrl + tab
Send ("993")
Send("{TAB 2}")
Send("s")
Send("{TAB 2}")
Send("t")
Send("+{TAB}");shift + tab - tabs backwards
Send("587")
Send ("{ENTER}")
Send ("n")
EndFunc

#comments-start
;Script Developed by Anthony Bickimer, OMCP
#comments-end