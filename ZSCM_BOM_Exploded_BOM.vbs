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
session.findById("wnd[0]").resizeWorkingPane 170,27,false
session.findById("wnd[0]/tbar[0]/okcd").text = "ZSCM_BOM"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_WERKS").text = "2040"
session.findById("wnd[0]/usr/ctxtP_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtP_WERKS").caretPosition = 4
session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
