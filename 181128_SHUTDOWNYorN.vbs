Set wshShell =wscript.CreateObject("WScript.Shell")
wscript.sleep 100

WshShell.Run "%windir%\notepad.exe"
wscript.sleep 100
WshShell.AppActivate "Notepad"
wscript.sleep 500

wshshell.sendkeys "{B}"
wscript.sleep 200
wshshell.sendkeys "{Y}"
wscript.sleep 300
wshshell.sendkeys "{E}"
wscript.sleep 200
wshshell.sendkeys "."
wscript.sleep 200
wshshell.sendkeys "."
wscript.sleep 100
wshshell.sendkeys "."
wscript.sleep 200


wshshell.sendkeys "%{F4}"
wscript.sleep 100
wshshell.sendkeys "{RIGHT}"
wscript.sleep 100
wshshell.sendkeys "{ENTER}"
wscript.sleep 100

Select Case MsgBox("종료할까요?", vbYesNo, "시스템")
Case vbYes
WshShell.Run "cmd.exe"
wscript.sleep 100
WshShell.AppActivate "Cmd"
wscript.sleep 100

       wshshell.sendkeys "{S}"
wscript.sleep 100
wshshell.sendkeys "{H}"
wscript.sleep 100
wshshell.sendkeys "{U}"
wscript.sleep 100
wshshell.sendkeys "{T}"
wscript.sleep 100
wshshell.sendkeys "{D}"
wscript.sleep 100
wshshell.sendkeys "{O}"
wscript.sleep 100
wshshell.sendkeys "{W}"
wscript.sleep 100
wshshell.sendkeys "{N}"
wscript.sleep 100
wshshell.sendkeys " "
wscript.sleep 100
wshshell.sendkeys "{-}"
wscript.sleep 100
wshshell.sendkeys "{S}"
wscript.sleep 100
wshshell.sendkeys "{G}"
wscript.sleep 200
wshshell.sendkeys " "
wscript.sleep 100
wshshell.sendkeys "{-}"
wscript.sleep 100
wshshell.sendkeys "{T}"
wscript.sleep 100
wshshell.sendkeys " "
wscript.sleep 100
wshshell.sendkeys "10"
wscript.sleep 200
wshshell.sendkeys "{ENTER}"
wscript.sleep 300
wshshell.sendkeys "{ESC}"
wscript.sleep 300
wshshell.sendkeys "%{F4}"
wscript.sleep 100



Select Case MsgBox("취소할까요?", vbYesNo, "시스템")
Case vbYes
WshShell.Run "cmd.exe"
wscript.sleep 100
WshShell.AppActivate "Cmd"
wscript.sleep 100

wshshell.sendkeys "{S}"
wscript.sleep 100
wshshell.sendkeys "{H}"
wscript.sleep 100
wshshell.sendkeys "{U}"
wscript.sleep 100
wshshell.sendkeys "{T}"
wscript.sleep 100
wshshell.sendkeys "{D}"
wscript.sleep 100
wshshell.sendkeys "{O}"
wscript.sleep 100
wshshell.sendkeys "{W}"
wscript.sleep 100
wshshell.sendkeys "{N}"
wscript.sleep 100
wshshell.sendkeys " "
wscript.sleep 100
wshshell.sendkeys "{-}"
wscript.sleep 100
wshshell.sendkeys "{A}"
wscript.sleep 100
wshshell.sendkeys "{ENTER}"
wscript.sleep 200




Case vbNo
MsgBox "종료"

End Select





Case vbNo
MsgBox "종료 취소됨"

End Select





