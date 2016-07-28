   Set oWS = WScript.CreateObject("WScript.Shell")
   sUserProfile = OWS.ExpandEnvironmentStrings("%USERPROFILE%")
   sWinDir = OWS.ExpandEnvironmentStrings("%SYSTEMROOT%")
   sLinkFile = sUserProfile & "\desktop\IT Help.LNK"
   Set oLink = oWS.CreateShortcut(sLinkFile)
   
   oLink.TargetPath = "https://eis.afsoc.af.mil/sites/413FLTS/CAX/CAXI/Lists/IT%20Help1/NewForm.aspx"

   	oLink.IconLocation = sWinDir & "\413th.ico"

   oLink.Save