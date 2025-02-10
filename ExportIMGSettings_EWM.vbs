Sub ExportIMGSettings_EWM()

'Extraction of EWM IMG settings based on conditions from an excel spreadsheet (with customizing nodes on pop-up screens)

Dim SapGui As Object
Dim App As SAPFEWSELib.GuiApplication
Dim Conn As SAPFEWSELib.GuiConnection
Dim Session As SAPFEWSELib.GuiSession

Dim Path As Range
Dim NodeName As Range
Dim Directory As Range
Dim Position As Range
Dim IDName As Range
Dim NodeKey As String
Dim CurrDate As String
Dim y As String
Const1 = "2" 'ID of an item in a line

Set SapGui = GetObject("SAPGUI")
If IsObject(SapGui) Then
    Set App = SapGui.GetScriptingEngine
    If IsObject(App) Then
        Set Conn = App.Children(0)
        If IsObject(Conn) Then
            Set Session = Conn.Children(0)
            
                'Get current date to name extracted files correctly
                
                CurrDate = Format(Date, "dd.mm.yyyy")
                
                'Start a t-code and open an IMG tree
                
                Session.FindById("wnd[0]").maximize
                Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nspro"
                Session.FindById("wnd[0]").sendVKey 0
                Session.FindById("wnd[0]/tbar[1]/btn[5]").press
                
                'Expand all EWM tree nodes (10 clicks are required (BAdIs are not expanded))
                
                For i = 1 To 10

                    Session.FindById("wnd[0]").maximize
                    Session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem "02  1     28", "TEXT"
                    Session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem "02  1     28", "TEXT"
                    Session.FindById("wnd[0]/tbar[1]/btn[6]").press

                Next i
                
                Set guitree = Session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell")
                
                'Get required data from an active sheet
                
                Set Path = ActiveWorkbook.ActiveSheet.Range("A2:A5")
                Set Position = ActiveWorkbook.ActiveSheet.Range("B2:B5")
                Set NodeName = ActiveWorkbook.ActiveSheet.Range("C2:C5")
                Set IDName = ActiveWorkbook.ActiveSheet.Range("D2:D5")
                Set Directory = ActiveWorkbook.ActiveSheet.Range("E2:E5")
                
                'Loop over data selected from an active sheet and perform spreadsheet generation
                
                For j = 1 To Path.Count
                
                        NodeKey = guitree.GetNodeKeyByPath(Path(j))
                
                        Session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").selectItem NodeKey, Const1
                        Session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem NodeKey, Const1
                        Session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").clickLink NodeKey, Const1
                        
                        'Specific logic in case if a table with additional IMG nodes was opened
                        
                        If Position(j) <> "" Then
                            
                            y = Position(j).Value
                        
                            Session.FindById("wnd[1]/usr/tblSAPLDSYHTCODES/txtI_TSTCT-TTEXT[1," + y + "]").SetFocus
                            Session.FindById("wnd[1]/usr/tblSAPLDSYHTCODES/txtI_TSTCT-TTEXT[1," + y + "]").caretPosition = 13
                            Session.FindById("wnd[1]/tbar[0]/btn[2]").press
                            Session.FindById("wnd[0]/mbar/menu[0]/menu[9]/menu[0]").Select
                            Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = Directory(j)
                            Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = IDName(j) + "_" + CurrDate + ".XLSX"
                            Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
                            Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                            Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                            Session.FindById("wnd[1]").Close
                            
                            y = ""
                            
                        Else
                        
                        Session.FindById("wnd[0]/mbar/menu[0]/menu[9]/menu[0]").Select
                        Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = Directory(j)
                        Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = NodeName(j) + "_" + CurrDate + ".XLSX"
                        Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
                        Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                        Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                        
                        End If
                        
                        NodeKey = ""
                
                Next
                                           
                'Exit to the main screen after all data is collected
                
                Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n"
                Session.FindById("wnd[0]").sendVKey 0
                Session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").SelectedNode = "Root"
                            
        End If
    End If
End If

End Sub



