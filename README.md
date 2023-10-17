SAP GUI Automation Script

This repository contains a VBS script designed to automate specific tasks within the SAP GUI.

The script interacts with the SAP GUI to perform a series of actions, such as establishing a connection, initiating a session, and manipulating GUI elements based on their IDs.

Features:

    Initialization: The script initializes the SAP GUI scripting engine, ensuring a connection is established and a session is initiated.

    Maximize SAP Window: The main SAP window is maximized for better visibility during automation.

    Tab Existence Check: A function (IsTabThere) checks if a specific tab or element exists in the SAP GUI by its ID.

    Automated Interactions: The FindByIdLoop function performs a series of actions on the SAP GUI based on two parameters. It interacts with various GUI elements, sets values, and selects checkboxes.

    Looped Interactions: The script contains a loop to check multiple lines in the SAP GUI and perform actions on them.

Usage

To use the script:

    Ensure you have the SAP GUI installed and the scripting engine enabled.
    Run the script using a VBS interpreter or through the Windows Script Host.
    The script will automatically perform the defined tasks in the SAP GUI, in Process Director - PD, while processing invoices.
   

Contributions

Feel free to fork this repository and make any changes or improvements. Pull requests are welcome!

Code Breakdown:

     ' Initialize the SAP GUI Scripting Engine
     If Not IsObject(application) Then
        ' Get the SAP GUI object
        Set SapGuiAuto  = GetObject("SAPGUI")
        ' Get the scripting engine from the SAP GUI object
        Set application = SapGuiAuto.GetScriptingEngine
     End If
     
     ' Establish a connection to the SAP application
     If Not IsObject(connection) Then
        Set connection = application.Children(0)
     End If
     
     ' Initiate a session with the SAP application
     If Not IsObject(session) Then
        Set session    = connection.Children(0)
     End If
     
     ' Connect the SAP session and application to Windows Script Host (WScript)
     If IsObject(WScript) Then
        WScript.ConnectObject session,     "on"
        WScript.ConnectObject application, "on"
     End If
     
     ' Maximize the main SAP window
     session.findById("wnd[0]").maximize
     
     ' Function to check if a specific tab or element exists in the SAP GUI by its ID
     Function IsTabThere(tabId)
         ' Suppress errors
         On Error Resume Next
         Dim tab
         Set tab = session.findById(tabId)
         If Err.Number = 0 Then
             IsTabThere = True
         Else
             IsTabThere = False
         End If
         ' Resume normal error handling
         On Error GoTo 0
     End Function

 It is important to emphasize that, the below code will enter data into the 8th column, which in my layout is the tax code. If your 8th column is not tax code, the script might not be able to enter data due to conditional formatting. 
     
     ' Function to perform a series of actions on the SAP GUI based on two parameters
     Function FindByIdLoop(param1, param2)
         ' Suppress errors for the entire function
         On Error Resume Next       
     
         ' Interact with the SAP GUI using the first pattern
         Dim id1
         id1 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-TAX_CODE[8,0]"
         session.findById(id1).text = "INSERT_TAX_CODE"
     
         ' Interact with the SAP GUI using the second pattern
         Dim id2
         id2 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:" & param2 & "/tabsG_STRIP_HDR/tabpTAB3"
         session.findById(id2).select
     
         ' Additional actions with replaced ID parts
         Dim id3, id4, id5
         id3 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:" & param2 & "/tabsG_STRIP_HDR/tabpTAB3/ssubSUB:/COCKPIT/SAPLDISPLAY46:0403/tbl/COCKPIT/SAPLDISPLAY46G_TC_TAX_DET/ctxt/COCKPIT/STAX_DISP-TAX_CODE[0,0]"
         id4 = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_HDR:/COCKPIT/SAPLDISPLAY46:" & param2 & "/tabsG_STRIP_HDR/tabpTAB3/ssubSUB:/COCKPIT/SAPLDISPLAY46:0403/chk/COCKPIT/SHDR_DISP-CALC_TAX_IND"
         id5 = id4
     
         session.findById(id3).text = "INSERT_TAX_CODE"
         session.findById(id4).setFocus
         session.findById(id5).selected = true
         
If multiple lines are given to fill, the below loop attempts to go through and enter data.
         
     
         ' Loop to check multiple lines and perform actions on them
         Dim j, attempts, tabId
         For j = 0 To 9
             attempts = 0
             Do
                 tabId = "wnd[0]/usr/subSUB_MAIN:/COCKPIT/SAPLDISPLAY46:" & param1 & "/subSUB_ITEM:/COCKPIT/SAPLDISPLAY46:0410/tbl/COCKPIT/SAPLDISPLAY46G_TC_ITEM_DET/ctxt/COCKPIT/SITEM_DISP-TAX_CODE[8," & j & "]"
                 If IsTabThere(tabId) Then
                     session.findById(tabId).text = "INSERT_TAX_CODE"
                     session.findById(tabId).setFocus
                     session.findById(tabId).caretPosition = 2
                     session.findById("wnd[0]").sendVKey 0
                     Exit Do
                 End If
                 attempts = attempts + 1
             Loop Until attempts >= 5
         Next
     
         ' Resume normal error handling at the end of the function
         On Error Goto 0
     End Function

This the main issue with PR1 is the bug that your control ID "SAPLDISPLAY46:XXXX" can change whenever you process any item. Therefore you will not know (or at least i have not yet found), what your current ID number is. Upon collecting some data, the system switches between the below two set of ID-s (0381 / 0387 - 405 is constant), thus is a workaround.    
     
     ' Call the FindByIdLoop function with different sets of parameters
     FindByIdLoop "0381", "0405"
     FindByIdLoop "0387", "0405"
