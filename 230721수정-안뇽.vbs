
Set wshShell =wscript.CreateObject("WScript.Shell")
wscript.sleep 100
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 100
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}"
wscript.sleep 100
wshshell.sendkeys "{SCROLLLOCK}"
wscript.sleep 100
wshshell.sendkeys "{SCROLLLOCK}"
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}"
wscript.sleep 100
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 100
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 100

WshShell.Run "%windir%\notepad.exe"
wscript.sleep 100
WshShell.AppActivate "Notepad"
wscript.sleep 100

wshshell.sendkeys "{H}"
wscript.sleep 200
wshshell.sendkeys "{E}"
wscript.sleep 300
wshshell.sendkeys "{L}"
wscript.sleep 200
wshshell.sendkeys "{L}"
wscript.sleep 200
wshshell.sendkeys "{O}"
wscript.sleep 100
wshshell.sendkeys "{?}"
wscript.sleep 200


wshshell.sendkeys "{?}"
wscript.sleep 500

wshshell.sendkeys "^S"
wscript.sleep 100
wshshell.sendkeys "{ENTER}"
wscript.sleep 900
wshshell.sendkeys "{T}"
wscript.sleep 400
wshshell.sendkeys "{o}"
wscript.sleep 400
wshshell.sendkeys "{.}"
wscript.sleep 400
wshshell.sendkeys "{Y}"
wscript.sleep 400
wshshell.sendkeys "{o}"
wscript.sleep 400
wshshell.sendkeys "{u}"
wscript.sleep 400
wshshell.sendkeys "{ENTER}"
wscript.sleep 900

Select Case MsgBox("메모장을 닫을까요?", vbYesNo, "메모장")
Case vbYes
wshshell.sendkeys "%{F4}"
Case vbNo
MsgBox "그럼 알아서 닫으시길..."
End Select