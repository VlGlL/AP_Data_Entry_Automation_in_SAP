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
session.findById("wnd[0]").maximize

Function FindByIdLoop(param1, param2)
    ' Suppress errors for the entire function
    On Error Resume Next

    ' For the first pattern
    Dim id1
    id1 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-TAX_CODE[8,0]"
    session.findById(id1).text = "p5"

    ' For the second pattern
    Dim id2
    id2 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:" & param2 & "/tabsG_STRIP_HDR/tabpTAB3"
    session.findById(id2).select

    ' Additional actions with replaced ID parts
    Dim id3, id4, id5
    id3 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:" & param2 & "/tabsG_STRIP_HDR/tabpTAB3/ssubSUB:/COCKPIT/SAPLDISPLAY46:0403/tbl/COCKPIT/SAPLDISPLAY46G_TC_TAX_DET/ctxt/COCKPIT/STAX_DISP-TAX_CODE[0,0]"
    id4 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:" & param2 & "/tabsG_STRIP_HDR/tabpTAB3/ssubSUB:/COCKPIT/SAPLDISPLAY46:0403/chk/COCKPIT/SHDR_DISP-CALC_TAX_IND"
    id5 = id4

    session.findById(id3).text = "p5"
    session.findById(id4).setFocus
    session.findById(id5).selected = true

    ' Resume normal error handling at the end of the function
    On Error Goto 0
End Function

' Example of how to call the function for both sets of parameters
FindByIdLoop "0381", "0405"
FindByIdLoop "0387", "0405"
