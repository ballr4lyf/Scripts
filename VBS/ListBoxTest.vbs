Option Explicit 
Dim aOpt, sPrt
aOpt = Array("\\mail\HP 4100 - Clinical - Clare",    _ 
             "\\mail\HP 4000 - Admissions",          _ 
             "\\medical\HP LaserJet 4000 - Medical", _ 
             "\\mail\HP 4000 - Front Desk",          _ 
             "\\mail\HP 4000 - Business Office",     _ 
             "\\mail\HP 4100 - Clinical - HRC",      _ 
             "\\mail\HP 4100 - Business Office",     _ 
             "\\victoria\HP 4500"                    _ 
            ) 
'wsh.echo "You selected:", 
sPrt = SelectBox("Select a default printer", aOpt)
wscript.echo sprt


Function SelectBox(sTitle, aOptions) 
Dim oIE, s, item 
  set oIE = CreateObject("InternetExplorer.Application") 
  With oIE 
    '.FullScreen = True ' remove if using IE 7+ 
    .ToolBar   = False : .RegisterAsDropTarget = False 
    .StatusBar = False : .Navigate("about:blank") 
    While .Busy : WScript.Sleep 100 : Wend 
    .width= 400 : .height=175 
    With .document 
      with .parentWindow.screen 
        oIE.left = (.availWidth-oIE.width)\2 
        oIE.top  = (.availheight-oIE.height)\2 
      End With ' ParentWindow 
      s = "<html><head><title>" & sTitle _ 
        & "</title></head><script language=vbs>bWait=true<" & "/script>" _ 
        & "<body bgColor=ghostwhite><center><b>" & sTitle & "<b><p>" _ 
        & "<select id=entries size=1 style='width:325px'>" _ 
        & "  <option selected>" & sTitle & "</option>" 
      For each item in aOptions 
        s = s & "  <option>" & item & "</option>" 
      Next 
      s = s & "  </select><p>" _ 
        & "<button id=but0 onclick='bWait=false'>OK</button>" _ 
        & "</center></body></html>" 
      .open 
      .WriteLn(s) 
      .close 
      Do until .ReadyState ="complete" : Wscript.Sleep 50 : Loop 
      With .body 
        .scroll="no" 
        .style.borderStyle = "outset" 
        .style.borderWidth = "3px" 
      End With ' Body 
      .all.entries.focus 
      oIE.Visible = True 
      CreateObject("Wscript.Shell").AppActivate sTitle 
      On Error Resume Next 
      While .ParentWindow.bWait 
        WScript.Sleep 100 
        if oIE.Visible Then SelectBox = "Aborted" 
        if Err.Number <> 0 Then Exit Function 
      Wend ' Wait 
      On Error Goto 0 
      With .ParentWindow.entries 
        SelectBox = .options(.selectedIndex).text 
      End With 
    End With ' document 
    .Visible = False 
  End With   ' IE 
End Function 


