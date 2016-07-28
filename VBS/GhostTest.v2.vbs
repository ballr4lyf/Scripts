Set OWS = wscript.CreateObject("WScript.Shell") 

i = 0

do while i < 5

OWS.Run "notepad" 
WScript.Sleep 100 
OWS.AppActivate "Notepad" 
WScript.Sleep 500 
  
OWS.sendkeys "%" 
OWS.sendkeys "O" 
OWS.sendkeys "F" 
OWS.sendkeys "{TAB}" 
OWS.sendkeys "{TAB}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{ENTER}" 
OWS.sendkeys "%" 
OWS.Sendkeys "{ }" 
WScript.Sleep 500 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{DOWN}" 
OWS.sendkeys "{x}" 
  
OWS.sendkeys "{G}" 
wscript.sleep 80 
OWS.sendkeys "{O}" 
wscript.sleep 80 
OWS.sendkeys "{ }" 
wscript.sleep 80 
OWS.sendkeys "{N}" 
wscript.sleep 80 
OWS.sendkeys "{O}" 
wscript.sleep 80 
OWS.sendkeys "{L}" 
wscript.sleep 80 
OWS.sendkeys "{E}" 
wscript.sleep 80 
OWS.sendkeys "{S}" 
wscript.sleep 80 
OWS.sendkeys "!" 
wscript.sleep 80 
OWS.sendkeys "!" 
wscript.sleep 80 
OWS.sendkeys "!" 
wscript.sleep 80 
OWS.sendkeys "!" 
wscript.sleep 5000 
OWS.SendKeys "%F" 
wscript.sleep 750 
OWS.SendKeys "{x}" 
wscript.sleep 750 
OWS.SendKeys "{TAB}" 
wscript.sleep 750 
OWS.SendKeys "{Enter}" 
wscript.sleep 750 

i = i + 1

loop 
