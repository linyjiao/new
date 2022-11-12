set ws=WScript.CreateObject("WScript.Shell")
Dim wmp
Set wmp = CreateObject("WMPlayer.OCX")
wmp.URL = "music.mp3"
ws.Run "rainbow.bat"
ws.Run "rainbow.bat"
ws.Run "rainbow_start.vbs"
ws.Run "key.vbs"
ws.Run "error.vbs"
ws.Run "error_2.vbs"
ws.Run "rainbow.vbs"
wscript.sleep 3500
ws.Run "error_2.vbs"
ws.Run "time.vbs"
Do Until wmp.playState = 1
wscript.Sleep 1000
Loop