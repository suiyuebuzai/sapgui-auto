If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session     = connection.Children(0)
End If
If Not IsObject(WScript) Then
   Set WScript = CreateObject("WScript.Shell")
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "1000001"
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
