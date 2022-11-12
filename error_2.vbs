Set ws = Wscript.CreateObject("Wscript.Shell")    
wscript.sleep 2000
ws.run("error_2.vbs")
msgbox "Error!!!",16+4096,"Windows error"

