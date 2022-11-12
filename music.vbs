Set ws = Wscript.CreateObject("Wscript.Shell")    
wscript.sleep 5000
ws.run("music.vbs")
Dim wmp
Set wmp = CreateObject("WMPlayer.OCX")
wmp.URL = "music.mp3"
Do Until wmp.playState = 1
wscript.Sleep 1000
Loop

