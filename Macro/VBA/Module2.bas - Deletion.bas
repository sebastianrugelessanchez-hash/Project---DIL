Attribute VB_Name = "Module2"
Sub VL06G()
Application.DisplayAlerts = False

   Set SapGuiAuto = GetObject("SAPGUI")
   Set App = SapGuiAuto.GetScriptingEngine
   Set Connection = App.Children(0)
   Set session = Connection.Children(0)
   
Sheets("MACRO").Select
fecha1 = Range("d6").Value
fecha2 = Range("E6").Value
fecha3 = Range("I6").Value
RUTA = ThisWorkbook.Path
FECHA = Format(Now, "mm.dd.yyyy")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl06g"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").Text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey 0

'exclusion CEM
session.findById("wnd[0]/usr/btn%_IT_SPART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT").Select
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT/ssubSCREEN_HEADER:SAPLALDB:3040/tblSAPLALDBINTERVAL_E/ctxtRSCSEL_255-ILOW_E[1,0]").Text = "07"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT/ssubSCREEN_HEADER:SAPLALDB:3040/tblSAPLALDBINTERVAL_E/ctxtRSCSEL_255-IHIGH_E[2,0]").Text = "12"
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").Text = fecha1
session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").Text = fecha2
session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").press

Sheets("MACRO").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy



session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

'condicion no hay tickets
If session.findById("wnd[0]/sbar").Text Like "No deliveries selected" Then
    nodelivery = 1
    GoTo No_deliveries:
End If


session.findById("wnd[0]/tbar[1]/btn[18]").press

'session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 1, "TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 0
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/usr/ctxtLIKP-WADAT_IST").Text = fecha3
 session.findById("wnd[1]/usr/ctxtLIKP-WADAT_IST").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[20]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 2, "LFDAT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = RUTA & "\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "VL06G " & FECHA & ".XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press


session.findById("wnd[0]/tbar[1]/btn[9]").press
session.findById("wnd[0]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").setCurrentCell 0, "T_MSG"
session.findById("wnd[0]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = RUTA & "\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ERROR LOG " & FECHA & ".XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press


session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

libro2 = "VL06G " & FECHA & ".XLSX"
libro3 = "INVENTORY " & FECHA & ".XLSX"
libro4 = "ERROR LOG " & FECHA & ".XLSX"

    Workbooks.Open Filename:= _
        RUTA & "\" & libro2


On Error Resume Next
    Windows(libro2).Activate
    ActiveSheet.Name = "VL06G"
    

    Range("S1").Select
    ActiveCell.FormulaR1C1 = "0.90718"
    Columns("Q:Q").Select
    Selection.AutoFilter
    ActiveSheet.Range("$Q$1:$Q$1000000").AutoFilter Field:=1, Criteria1:="TO"
    Range("P1000000").Select
    Selection.End(xlUp).Select

i = ActiveCell.Row

If i <> 1 Then


    Range("S1").Select
    Selection.Copy
    Range("P2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "TON"
    Range("S1").Select
    Selection.ClearContents
    Selection.AutoFilter


    Columns("P:P").Select
    Selection.NumberFormat = "0.000"
    
End If
    
    
    
    Range("C39").Select
    Sheets("Sheet1").Select
    Sheets("VL06G").Copy After:=Workbooks("Inventory " & FECHA & ".xlsx").Sheets(1)
    Windows(libro3).Activate
    Range("B1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=3, Criteria1:="="

    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    Range("A1").Select
ActiveSheet.Name = "VL06G"


    Sheets("ERROR LOG").Select
    Columns("E:E").Select
    Calculate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=LEN(RC[-1])"
    Range("D1").Select
    Selection.End(xlDown).Select
l = ActiveCell.Row
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & l)
    Columns("E:E").Select
    Selection.AutoFilter
    ActiveSheet.Range("$E$1:$E$" & l).AutoFilter Field:=1, Criteria1:="<=3", _
        Operator:=xlAnd
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select





On Error GoTo SIGUEEE1: On Error GoTo -1
    Sheets("INVENTORY").Select
    Range("A1").Select
    Selection.Copy
    Sheets("VL06G").Select
    Range("A1:P1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
On Error GoTo 0
SIGUEEE1:
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
'aca poner la hoja en la macro

    Cells.Select
    Selection.Copy
    Windows("AUDIT MACRO.xlsm").Activate
    Sheets("VL06G").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Windows("AUDIT MACRO.xlsm").Activate
    
    Workbooks.Open Filename:= _
        RUTA & "\" & libro4

    Windows(libro4).Activate
    Sheets("Sheet1").Select
On Error GoTo MONDA: On Error GoTo -1
    Sheets("Sheet1").Copy After:=Workbooks(libro3).Sheets(1)
On Error GoTo 0
VOLVERR1:
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "ERROR LOG"
    
'ACA PONER EL FILTRO

On Error GoTo MONDA2: On Error GoTo -1
    Windows(libro3).Activate
On Error GoTo 0
VOLVERRRR1:
    Range("E1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$F$1000000").AutoFilter Field:=5, Criteria1:= _
        "=*DEFICIT*", Operator:=xlAnd
    ActiveWindow.SmallScroll Down:=-9
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    On Error GoTo sigueeee: On Error GoTo -1
    ActiveSheet.ShowAllData
    On Error GoTo 0
sigueeee:
    Selection.AutoFilter
    Windows("AUDIT MACRO.xlsm").Activate
    Sheets("AUDIT").Select
    Range("I1").Select
    Selection.Copy
On Error GoTo MONDA3: On Error GoTo -1
    Windows(libro3).Activate
On Error GoTo 0
VOLVERRRR2:
    Range("A1:F1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C13").Select
    Sheets("VL06G").Select
    Range("E1").Select
    Sheets("ERROR LOG").Select
    Range("C1").Select
    Selection.Copy
    Sheets("VL06G").Select
    Range("A1:P1").Select
    Range("P1").Activate
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A2").Select
    Sheets("ERROR LOG").Select

    Sheets("VL06G").Select
    Range("P1").Select
    Selection.End(xlDown).Select
l = ActiveCell.Row
    Selection.End(xlUp).Select
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-16],'ERROR LOG'!C[-14],1,0)"
    Selection.AutoFill Destination:=Range("Q2:Q" & l)
    Range("Q2:Q79").Select
    Columns("Q:Q").Select
    Selection.AutoFilter
    Range("Q4").Select
    ActiveSheet.Range("$Q$1:$Q$1000000").AutoFilter Field:=1, Criteria1:="<>#N/A", _
        Operator:=xlAnd
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    Columns("Q:Q").Select
    Selection.ClearContents
    Range("A1").Select


    
    'Windows(libro2).Activate
    'ActiveWindow.Close
    Windows("AUDIT MACRO.xlsm").Activate
        Sheets("PIVOT INV").Select

'refrescar tabla y crear el archivo con los nuevos datos

    Windows("AUDIT MACRO.xlsm").Activate
    Windows("VL06G " & FECHA & ".XLSX").Activate
    Sheets("ERROR LOG").Select
    Cells.Select
    Selection.Copy
    Windows("AUDIT MACRO.xlsm").Activate
    Sheets("ERROR LOG").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C7").Select


Sheets("vl06g").Select
    Range("S1").Select
    Selection.ClearContents


    Sheets("PIVOT INV").Select
    Range("C8").Select
    Application.CutCopyMode = False
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    ActiveWindow.SmallScroll Down:=3
    Sheets(Array("ERROR LOG", "VL06G", "PIVOT INV")).Select
    Sheets("ERROR LOG").Activate
    Sheets(Array("ERROR LOG", "VL06G", "PIVOT INV")).Copy
    Sheets("PIVOT INV").Select
 
    
        ActiveWorkbook.SaveAs Filename:=RUTA & "\VL06G FINAL " & FECHA & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    'ActiveWindow.Close
    
Call CerraLibros
    
Exit Sub

MONDA:
    Sheets("Sheet1").Copy After:=Workbooks(libro2).Sheets(1)
GoTo VOLVERR1

MONDA2:
    Windows(libro2).Activate
GoTo VOLVERRRR1

MONDA3:
    Windows(libro2).Activate
GoTo VOLVERRRR2

If nodelivery = 1 Then
No_deliveries:
    MsgBox "No deleveries were selected so no PGI report was created"
End If

End Sub

