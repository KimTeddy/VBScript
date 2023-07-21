Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
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



if colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count -1
colCDROMs.Item(i).Eject

Next
End If
wscript.sleep 5000
