If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]/usr/subSUB7:ZBARCODE_SCANNER_COPY:0908/radR_ST20").setFocus
session.findById("wnd[0]/usr/subSUB7:ZBARCODE_SCANNER_COPY:0908/radR_ST20").select
session.findById("wnd[0]").sendVKey 8
